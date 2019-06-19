VERSION 5.00
Begin VB.Form Frm_Cat_Gaps_Layout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GAPS"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Imprimir 
      Caption         =   "Imprimir"
      Height          =   285
      Left            =   10335
      TabIndex        =   28
      Top             =   75
      Width           =   1080
   End
   Begin VB.Frame Fra_Gaps 
      BackColor       =   &H00FFFFFF&
      Height          =   8760
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11325
      Begin VB.TextBox Txt_Ruta_Imagen_9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8595
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   7230
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   7230
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4665
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   7230
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8595
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4665
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4665
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8595
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2070
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2070
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_Ruta_Imagen_1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2070
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image Img_Foto_9 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   8580
         Picture         =   "Frm_Cat_Gaps_Layout.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   6120
         Width           =   1725
      End
      Begin VB.Label Lbl_Empleado_9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   7957
         TabIndex        =   31
         Top             =   7665
         Width           =   2970
      End
      Begin VB.Label Lbl_Puesto_9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7957
         TabIndex        =   30
         Top             =   8340
         Width           =   2970
      End
      Begin VB.Label Lbl_Comentarios_Gap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "GAP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5295
         TabIndex        =   27
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Lbl_Nombre_Gap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "GAP"
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
         Left            =   5220
         TabIndex        =   26
         Top             =   165
         Width           =   810
      End
      Begin VB.Label Lbl_Puesto_7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   25
         Top             =   8340
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   330
         TabIndex        =   24
         Top             =   7665
         Width           =   2970
      End
      Begin VB.Image Img_Foto_7 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   960
         Picture         =   "Frm_Cat_Gaps_Layout.frx":C042
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   6120
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   22
         Top             =   5775
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   330
         TabIndex        =   21
         Top             =   5070
         Width           =   2970
      End
      Begin VB.Image Img_Foto_4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   960
         Picture         =   "Frm_Cat_Gaps_Layout.frx":18084
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4132
         TabIndex        =   19
         Top             =   8340
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   4132
         TabIndex        =   18
         Top             =   7665
         Width           =   2970
      End
      Begin VB.Image Img_Foto_8 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   4755
         Picture         =   "Frm_Cat_Gaps_Layout.frx":240C6
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   6120
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7957
         TabIndex        =   16
         Top             =   5775
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   7950
         TabIndex        =   15
         Top             =   5070
         Width           =   2970
      End
      Begin VB.Image Img_Foto_6 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   8580
         Picture         =   "Frm_Cat_Gaps_Layout.frx":30108
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4132
         TabIndex        =   13
         Top             =   5775
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   4125
         TabIndex        =   12
         Top             =   5070
         Width           =   2970
      End
      Begin VB.Image Img_Foto_5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   4755
         Picture         =   "Frm_Cat_Gaps_Layout.frx":3C14A
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7957
         TabIndex        =   10
         Top             =   3225
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   7957
         TabIndex        =   9
         Top             =   2520
         Width           =   2970
      End
      Begin VB.Image Img_Foto_3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   8580
         Picture         =   "Frm_Cat_Gaps_Layout.frx":4818C
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4132
         TabIndex        =   7
         Top             =   3225
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   4132
         TabIndex        =   6
         Top             =   2520
         Width           =   2970
      End
      Begin VB.Image Img_Foto_2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   4755
         Picture         =   "Frm_Cat_Gaps_Layout.frx":541CE
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Lbl_Puesto_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Puesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   4
         Top             =   3225
         Width           =   2970
      End
      Begin VB.Label Lbl_Empleado_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   330
         TabIndex        =   3
         Top             =   2520
         Width           =   2970
      End
      Begin VB.Image Img_Foto_1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1425
         Left            =   960
         Picture         =   "Frm_Cat_Gaps_Layout.frx":60210
         Stretch         =   -1  'True
         ToolTipText     =   "Doble click para cambiar la imagen"
         Top             =   960
         Width           =   1725
      End
   End
   Begin VB.Image Img_Logo_Empresa 
      Height          =   735
      Left            =   120
      Picture         =   "Frm_Cat_Gaps_Layout.frx":6C252
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Lbl_Gaps 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TRIPULACION DE EMPLEADOS"
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
      Left            =   3015
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
End
Attribute VB_Name = "Frm_Cat_Gaps_Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'NOMBRE_FUNCION: Consulta_Empleados_Gap
'DESCRIPCION: Consulta los empleados que pertenecen al GAP y los muestra en pantalla
'PARAMETROS : Gap_ID- Indentificador del Gap que será mostrado en pantalla
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 13-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Public Sub Consulta_Empleados_Gap(Gap_ID As String)
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Contador_Empleado As Integer
    
    Limpia_Controles
    'Consulta los empleados activos que pertenecen al gap consultado
    Mi_SQL = "SELECT Cat_Empleados.Empleado_ID,Cat_Empleados.No_Tarjeta,Cat_Empleados.Apellido_Paterno+' '+Cat_Empleados.Apellido_Materno+' '+Cat_Empleados.Nombre AS Nombre_Empleado,Cat_Empleados.Imagen_Perfil,Cat_Puestos.Nombre AS Puesto"
    Mi_SQL = Mi_SQL & " FROM Cat_Empleados,Cat_Puestos"
    Mi_SQL = Mi_SQL & " WHERE Cat_Empleados.Puesto_ID=Cat_Puestos.Puesto_ID"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Gap_ID='" & Trim(Gap_ID) & "'"
    Mi_SQL = Mi_SQL & " AND Cat_Empleados.Estatus='A'"
    Mi_SQL = Mi_SQL & " ORDER BY Cat_Empleados.No_Tarjeta"
    Set Rs_Consulta_Cat_Empleados = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    While Not Rs_Consulta_Cat_Empleados.EOF
        With Rs_Consulta_Cat_Empleados
            Contador = Contador + 1
            Select Case Contador
                Case 1
                    Lbl_Empleado_1.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_1.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_1.Text = ""
                    Img_Foto_1.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_1.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_1.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_1.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 2
                    Lbl_Empleado_2.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_2.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_2.Text = ""
                    Img_Foto_2.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_2.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_2.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_2.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 3
                    Lbl_Empleado_3.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_3.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_3.Text = ""
                    Img_Foto_3.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_3.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_3.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_3.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 4
                    Lbl_Empleado_4.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_4.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_4.Text = ""
                    Img_Foto_4.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_4.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_4.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_4.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 5
                    Lbl_Empleado_5.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_5.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_5.Text = ""
                    Img_Foto_5.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_5.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_5.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_5.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 6
                    Lbl_Empleado_6.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_6.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_6.Text = ""
                    Img_Foto_6.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_6.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_6.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_6.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 7
                    Lbl_Empleado_7.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_7.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_7.Text = ""
                    Img_Foto_7.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_7.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_7.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_7.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 8
                    Lbl_Empleado_8.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_8.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_8.Text = ""
                    Img_Foto_8.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_8.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_8.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_8.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
                Case 9
                    Lbl_Empleado_9.Caption = .rdoColumns("No_Tarjeta") & " - " & .rdoColumns("Nombre_Empleado")
                    Lbl_Puesto_9.Caption = .rdoColumns("Puesto")
                    Txt_Ruta_Imagen_9.Text = ""
                    Img_Foto_9.picture = LoadPicture("")
                    If .rdoColumns("Imagen_Perfil") <> "" Then
                        'Valida que el archivo exista
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"), "ARCHIVO") = True Then
                            Img_Foto_9.picture = LoadPicture(App.Path & "\Perfil\" & .rdoColumns("Imagen_Perfil"))
                        Else
                            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                                Img_Foto_9.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                            End If
                        End If
                    Else
                        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(App.Path & "\Perfil\User.bmp", "ARCHIVO") = True Then
                            Img_Foto_9.picture = LoadPicture(App.Path & "\Perfil\User.bmp")
                        End If
                    End If
            End Select
            .MoveNext
        End With
    Wend
    Rs_Consulta_Cat_Empleados.Close
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Limpia_Controles
'DESCRIPCION: Limpia los controles de la pantalla
'PARAMETROS :
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO : 13-Marzo-2012
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'******************************************************************************
Private Sub Limpia_Controles()
Dim Rs_Consulta_Cat_Empleados As rdoResultset
Dim Contador_Empleado As Integer
    
    Lbl_Empleado_1.Caption = ""
    Lbl_Puesto_1.Caption = ""
    Txt_Ruta_Imagen_1.Text = ""
    Img_Foto_1.picture = LoadPicture("")
    Lbl_Empleado_2.Caption = ""
    Lbl_Puesto_2.Caption = ""
    Txt_Ruta_Imagen_2.Text = ""
    Img_Foto_2.picture = LoadPicture("")
    Lbl_Empleado_3.Caption = ""
    Lbl_Puesto_3.Caption = ""
    Txt_Ruta_Imagen_3.Text = ""
    Img_Foto_3.picture = LoadPicture("")
    Lbl_Empleado_4.Caption = ""
    Lbl_Puesto_4.Caption = ""
    Txt_Ruta_Imagen_4.Text = ""
    Img_Foto_4.picture = LoadPicture("")
    Lbl_Empleado_5.Caption = ""
    Lbl_Puesto_5.Caption = ""
    Txt_Ruta_Imagen_5.Text = ""
    Img_Foto_5.picture = LoadPicture("")
    Lbl_Empleado_6.Caption = ""
    Lbl_Puesto_6.Caption = ""
    Txt_Ruta_Imagen_6.Text = ""
    Img_Foto_6.picture = LoadPicture("")
    Lbl_Empleado_7.Caption = ""
    Lbl_Puesto_7.Caption = ""
    Txt_Ruta_Imagen_7.Text = ""
    Img_Foto_7.picture = LoadPicture("")
    Lbl_Empleado_8.Caption = ""
    Lbl_Puesto_8.Caption = ""
    Txt_Ruta_Imagen_8.Text = ""
    Img_Foto_8.picture = LoadPicture("")
    Lbl_Empleado_9.Caption = ""
    Lbl_Puesto_9.Caption = ""
    Txt_Ruta_Imagen_9.Text = ""
    Img_Foto_9.picture = LoadPicture("")
End Sub

Private Sub Btn_Imprimir_Click()
Dim Mi_Impresora As Printer
On Error GoTo handler
    'Imprime la forma
    Btn_Imprimir.Visible = False
    MDIFrm_Apl_Principal.CommonDialog1.ShowPrinter
    'PrintForm
    'Comienza la impresion del encabezado
    Printer.ScaleMode = vbCentimeters
    'Printer.Orientation = vbHorizontal
    Printer.FontSize = 10
    Printer.Font = "Arial"
    Printer.FontBold = True
    Call Printer.PaintPicture(Img_Logo_Empresa.picture, 1, 0.5, 6, 2)
    Printer.FontSize = 16
    Printer.CurrentX = 8
    Printer.CurrentY = 0.5
    Printer.Print "TRIPULACION DE EMPLEADOS"
    Printer.CurrentX = 8
    Printer.CurrentY = 1.2
    Printer.Print Lbl_Nombre_Gap.Caption
    Printer.CurrentX = 8
    Printer.CurrentY = 1.9
    Printer.Print Lbl_Comentarios_Gap.Caption
    Printer.Line (0.5, 0.25)-(27.5, 21.25), , B
    'Imagenes de empleados
    Printer.FontSize = 10
    Printer.FontBold = False
    If Lbl_Empleado_1.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_1.picture, 1, 3, 4, 4)
        Printer.CurrentX = 1
        Printer.CurrentY = 7.3
        Printer.Print Lbl_Empleado_1.Caption
        Printer.CurrentX = 1
        Printer.CurrentY = 7.6
        Printer.Print Lbl_Puesto_1.Caption
    End If
    If Lbl_Empleado_2.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_2.picture, 10, 3, 4, 4)
        Printer.CurrentX = 10
        Printer.CurrentY = 7.3
        Printer.Print Lbl_Empleado_2.Caption
        Printer.CurrentX = 10
        Printer.CurrentY = 7.6
        Printer.Print Lbl_Puesto_2.Caption
    End If
    If Lbl_Empleado_3.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_3.picture, 19, 3, 4, 4)
        Printer.CurrentX = 19
        Printer.CurrentY = 7.3
        Printer.Print Lbl_Empleado_3.Caption
        Printer.CurrentX = 19
        Printer.CurrentY = 7.6
        Printer.Print Lbl_Puesto_3.Caption
    End If
    If Lbl_Empleado_4.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_4.picture, 1, 9, 4, 4)
        Printer.CurrentX = 1
        Printer.CurrentY = 13.3
        Printer.Print Lbl_Empleado_4.Caption
        Printer.CurrentX = 1
        Printer.CurrentY = 13.6
        Printer.Print Lbl_Puesto_4.Caption
    End If
    If Lbl_Empleado_5.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_5.picture, 10, 9, 4, 4)
        Printer.CurrentX = 10
        Printer.CurrentY = 13.3
        Printer.Print Lbl_Empleado_5.Caption
        Printer.CurrentX = 10
        Printer.CurrentY = 13.6
        Printer.Print Lbl_Puesto_5.Caption
    End If
    If Lbl_Empleado_6.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_6.picture, 19, 9, 4, 4)
        Printer.CurrentX = 19
        Printer.CurrentY = 13.3
        Printer.Print Lbl_Empleado_6.Caption
        Printer.CurrentX = 19
        Printer.CurrentY = 13.6
        Printer.Print Lbl_Puesto_6.Caption
    End If
    If Lbl_Empleado_7.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_7.picture, 1, 15, 4, 4)
        Printer.CurrentX = 1
        Printer.CurrentY = 19.3
        Printer.Print Lbl_Empleado_7.Caption
        Printer.CurrentX = 1
        Printer.CurrentY = 19.6
        Printer.Print Lbl_Puesto_7.Caption
    End If
    If Lbl_Empleado_8.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_8.picture, 10, 15, 4, 4)
        Printer.CurrentX = 10
        Printer.CurrentY = 19.3
        Printer.Print Lbl_Empleado_8.Caption
        Printer.CurrentX = 10
        Printer.CurrentY = 19.6
        Printer.Print Lbl_Puesto_8.Caption
    End If
    If Lbl_Empleado_9.Caption <> "" Then
        Call Printer.PaintPicture(Img_Foto_9.picture, 19, 15, 4, 4)
        Printer.CurrentX = 19
        Printer.CurrentY = 19.3
        Printer.Print Lbl_Empleado_9.Caption
        Printer.CurrentX = 19
        Printer.CurrentY = 19.6
        Printer.Print Lbl_Puesto_9.Caption
    End If
    Printer.EndDoc
    MsgBox "Tripulación enviada a impresion", vbInformation
    Btn_Imprimir.Visible = True
Exit Sub
handler:
    MsgBox "Ha ocurrido un problema al enviar a impresión, verifique su impresora esté encendida y funcionando correctamte", vbExclamation
    Btn_Imprimir.Visible = True
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Picture1_Click()

End Sub
