VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Apl_Configuracion_Checador 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8820
   Begin zkemkeeperCtl.CZKEM Checador 
      Height          =   105
      Left            =   7995
      OleObjectBlob   =   "Frm_Apl_Configuracion_Checador.frx":0000
      TabIndex        =   49
      Top             =   345
      Width           =   165
   End
   Begin VB.CommandButton Btn_Cerrar_Conexion 
      Caption         =   "Cerrar Conexion"
      Height          =   315
      Left            =   6240
      TabIndex        =   42
      Top             =   600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog Cmd_Dispositivos 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Cmb_Dipositivos 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin TabDlg.SSTab Tab_Dipositivo 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Informacion Dispositivo"
      TabPicture(0)   =   "Frm_Apl_Configuracion_Checador.frx":0024
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Informacion_Dispositivo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Configuracion"
      TabPicture(1)   =   "Frm_Apl_Configuracion_Checador.frx":0040
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Configuracion_Equipo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Administracion"
      TabPicture(2)   =   "Frm_Apl_Configuracion_Checador.frx":005C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Mantenimiento"
      TabPicture(3)   =   "Frm_Apl_Configuracion_Checador.frx":0078
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   2955
         Left            =   -75000
         TabIndex        =   43
         Top             =   360
         Width           =   8715
         Begin VB.Frame Frm_Mantenimiento 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Equipo"
            Height          =   1815
            Left            =   2520
            TabIndex        =   47
            Top             =   240
            Width           =   2235
            Begin VB.CommandButton Btn_Apagar_Equipo 
               Caption         =   "Apagar"
               Height          =   315
               Left            =   60
               TabIndex        =   37
               Top             =   300
               Width           =   1755
            End
            Begin VB.CommandButton Btn_Reinicar 
               Caption         =   "Reiniciar"
               Height          =   315
               Left            =   60
               TabIndex        =   48
               Top             =   1020
               Width           =   1755
            End
         End
         Begin VB.Frame Frm_Registros 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Registros"
            Height          =   1815
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2235
            Begin VB.CommandButton Btn_Limpiar_Información 
               Caption         =   "Limpiar Información"
               Height          =   315
               Left            =   60
               TabIndex        =   46
               Top             =   1020
               Width           =   1755
            End
            Begin VB.CommandButton Btn_LimpiarGLog 
               Caption         =   "Limpiar Registros"
               Height          =   315
               Left            =   60
               TabIndex        =   45
               Top             =   300
               Width           =   1755
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   2955
         Left            =   -74940
         TabIndex        =   38
         Top             =   360
         Width           =   8715
         Begin VB.CommandButton Btn_Excel 
            Caption         =   "Exportar Excel"
            Height          =   315
            Left            =   4597
            TabIndex        =   51
            Top             =   135
            Width           =   1755
         End
         Begin VB.CommandButton Btn_Obtener_Usuarios 
            Caption         =   "Obtener Usuarios"
            Height          =   315
            Left            =   2355
            TabIndex        =   50
            Top             =   135
            Width           =   1755
         End
         Begin VB.CommandButton Btn_Limpiar_Lista 
            Caption         =   "Limpiar Lista"
            Height          =   315
            Left            =   6840
            TabIndex        =   52
            Top             =   135
            Width           =   1755
         End
         Begin VB.Frame Fra_Dispositivo_Usuarios 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Usuarios"
            Height          =   2475
            Left            =   60
            TabIndex        =   39
            Top             =   405
            Width           =   8580
            Begin MSFlexGridLib.MSFlexGrid Grid_Usuarios 
               Height          =   2130
               Left            =   90
               TabIndex        =   40
               Top             =   255
               Width           =   8385
               _ExtentX        =   14790
               _ExtentY        =   3757
               _Version        =   393216
               Rows            =   0
               FixedRows       =   0
               BackColorBkg    =   16777215
               AllowUserResizing=   1
               Appearance      =   0
            End
         End
      End
      Begin VB.Frame Fra_Informacion_Dispositivo 
         BackColor       =   &H00FFFFFF&
         Height          =   2955
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   8715
         Begin VB.Frame Fra_Informacion_Dispositivo_Generales 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Generales"
            Height          =   2415
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   4095
            Begin VB.TextBox Txt_Formato_Fecha 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   180
               Width           =   2475
            End
            Begin VB.TextBox Txt_Mac 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   900
               Width           =   2475
            End
            Begin VB.CommandButton Btn_Actualizar_Firmware 
               Caption         =   "Actualizar Firmware"
               Height          =   315
               Left            =   1380
               TabIndex        =   29
               Top             =   2040
               Width           =   2475
            End
            Begin VB.TextBox Txt_Firmware 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   1620
               Width           =   2475
            End
            Begin VB.TextBox Txt_Numero_Serie 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   1260
               Width           =   2475
            End
            Begin VB.TextBox Txt_Direccion_IP 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   540
               Width           =   2475
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Formato Fecha"
               Height          =   195
               Left            =   60
               TabIndex        =   36
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Direccion MAC"
               Height          =   195
               Left            =   60
               TabIndex        =   34
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Firmware"
               Height          =   195
               Left            =   60
               TabIndex        =   32
               Top             =   1680
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Numero de Serie"
               Height          =   195
               Left            =   60
               TabIndex        =   31
               Top             =   1320
               Width           =   1185
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Direccion IP"
               Height          =   195
               Left            =   60
               TabIndex        =   30
               Top             =   600
               Width           =   870
            End
         End
         Begin VB.Frame Fra_Informacion_Usuarios 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Informacion de Usuarios"
            Height          =   2415
            Left            =   4440
            TabIndex        =   12
            Top             =   300
            Width           =   4095
            Begin VB.TextBox Txt_Transacciones 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox Txt_Logs_Administrador 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox Txt_No_Contraseñas 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox Txt_No_Plantilla 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox Txt_No_Usuarios 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox Txt_No_Administradores 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2700
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transacciones"
               Height          =   195
               Left            =   240
               TabIndex        =   24
               Top             =   2100
               Width           =   1050
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Logs de Administrador"
               Height          =   195
               Left            =   240
               TabIndex        =   23
               Top             =   1740
               Width           =   1560
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contraseñas"
               Height          =   195
               Left            =   240
               TabIndex        =   22
               Top             =   1380
               Width           =   885
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plantillas"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Usuarios"
               Height          =   195
               Left            =   240
               TabIndex        =   20
               Top             =   660
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Administradores"
               Height          =   195
               Left            =   240
               TabIndex        =   19
               Top             =   300
               Width           =   1110
            End
         End
      End
      Begin VB.Frame Fra_Configuracion_Equipo 
         BackColor       =   &H00FFFFFF&
         Height          =   2955
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   8715
         Begin VB.Frame Fra_Dispositivo_Fecha 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Configuracion Fecha"
            Height          =   2715
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   4275
            Begin VB.CommandButton Btn_Obtener_Fecha 
               Caption         =   "Obtener Fecha"
               Height          =   315
               Left            =   180
               TabIndex        =   8
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton Btn_Colocar_Fecha 
               Caption         =   "Colocar Fecha"
               Height          =   315
               Left            =   180
               TabIndex        =   7
               Top             =   810
               Width           =   1335
            End
            Begin VB.ComboBox Cmb_Formato_Fechas 
               Height          =   315
               ItemData        =   "Frm_Apl_Configuracion_Checador.frx":0094
               Left            =   2100
               List            =   "Frm_Apl_Configuracion_Checador.frx":00B6
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1260
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.CommandButton Btn_Cambiar_Formato 
               Caption         =   "Cambiar Formato"
               Enabled         =   0   'False
               Height          =   315
               Left            =   180
               TabIndex        =   5
               Top             =   1260
               Visible         =   0   'False
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Dispositivo 
               Height          =   315
               Left            =   2100
               TabIndex        =   9
               Top             =   360
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd MMM yyyy HH:mm:ss"
               Format          =   112787459
               UpDown          =   -1  'True
               CurrentDate     =   40156
            End
            Begin MSComCtl2.DTPicker Dtp_Fecha_Dispositivo1 
               Height          =   315
               Left            =   2100
               TabIndex        =   10
               Top             =   810
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd MMM yyyy HH:mm:ss"
               Format          =   112787459
               UpDown          =   -1  'True
               CurrentDate     =   40156
            End
         End
      End
   End
   Begin VB.Label LblEstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconectado"
      Height          =   195
      Left            =   1080
      TabIndex        =   41
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivo"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "Frm_Apl_Configuracion_Checador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IP As String
Dim Puerto As Long
Dim Maquina As Long
Public Equipo_ID As String
'Enumeraciones para la informacion de los dispositivos
Public Enum Informacion_Checador_Idiomas
    Checador_Informacion_Idioma_Ingles = 0
    Checador_Informacion_Idioma_Chino = 1
    Checador_Informacion_Idioma_Koreano = 2
End Enum
Public Enum Informacion_Checador_Paridad
    Checador_Informacion_Paridad_No = 0
    Checador_Informacion_Paridad_Par = 1
    Checador_Informacion_Paridad_Impar = 2
End Enum
Public Enum Informacion_Checador_Conexion_Red
    Checador_Informacion_Conexion_Red_Activa = 1
    Checador_Informacion_Conexion_Red_InActiva = 0
End Enum
Public Enum Informacion_Checador_Conexion_RS232
    Checador_Informacion_Conexion_RS232_Activa = 1
    Checador_Informacion_Conexion_RS232_InActiva = 0
End Enum
Public Enum Informacion_Checador_Conexion_RS485
    Checador_Informacion_Conexion_RS485_Activa = 1
    Checador_Informacion_Conexion_RS485_InActiva = 0
End Enum
Public Enum Informacion_Checador_Registro_Voz
    Checador_Informacion_Registro_Voz_Activa = 1
    Checador_Informacion_Registro_Voz_InActiva = 0
End Enum
Public Enum Informacion_Checador_Marcador
    Checador_Informacion_Marcador_Mostrar = 1
    Checador_Informacion_Marcador_Ocultar = 0
End Enum
Public Enum Informacion_Checador_Tarjeta_Verificacion
    Checador_Informacion_Tarjeta_Verificacion_SI = 1
    Checador_Informacion_Tarjeta_Verificacion_NO = 0
End Enum
Public Enum Checador_Informacion_Debera_Registrar_Tarjeta
    Checador_Informacion_Debera_Registrar_Tarjeta_SI = 1
    Checador_Informacion_Debera_Registrar_Tarjeta_NO = 0
End Enum
Public Enum Checador_Informacion
    Checador_Informacion_No_Maximo_Administradores = 1
    Checador_Informacion_No_Dispositivo = 2
    Checador_Informacion_Idioma = 3
    Checador_Informacion_Cerra_Dispositivo_Tiempo = 4
    Checador_Informacion_Exportar_Senal_Mentor = 5
    Checador_Informacion_No_Maximo_Atencion_Alarmas = 6
    Checador_Informacion_No_Maximo_Administracion_Noticias_Alarmas = 7
    Checador_Informacion_Espacio_Min_Validacion_Dos = 8
    Checador_Informacion_Tasa_Transferencia = 9
    'Diferencte de 6 se calcula 1200 * (dwValue + 1) & "bps"
    'si es & Rate:115200bps"
    Checador_Informacion_Paridad_Transferencia = 10
    Checador_Informacion_Senal_Detener = 11
    'Calculo (dwValue + 1) * 2 & "Bits"
    Checador_Informacion_Segmentacion_Informacion = 12
    'valor dwvalue "Data segmentation sign:" / ""
    'valor dwvalue "Data segmentation sign:-"
    Checador_Informacion_Conexion_Red = 13
    Checador_Informacion_Conexion_RS232 = 14
    Checador_Informacion_Conexion_RS485 = 15
    Checador_Informacion_Registro_Voz = 16
    Checador_Informacion_Valida_Velocidad = 17
    Checador_Informacion_Tiempo_Parado = 18
    Checador_Informacion_Tiempo_Cerrado = 19
    Checador_Informacion_Tiempo_Configuracion = 20
    Checador_Informacion_Tiempo_Dormir = 21
    Checador_Informacion_Beep_Operacion = 22
    Checador_Informacion_Minimo_Coincidencia = 23
    Checador_Informacion_Minimo_No_Coincidencia = 24
'    Checador_Informacion_Minimi_No_Coincidencia = 25
'        If I = 25 Then
'            If dwValue = 1 Then
'                ls2.AddItem "1:1:YES"
'                ls2.Refresh
'            Else
'                ls2.AddItem "1:1:NO"
'                ls2.Refresh
'            End If
'        End If
    Checador_Informacion_Mostrar_Marcador = 26
    Checador_Informacion_Combinacion_Desbloqueo_Personas = 27
    Checador_Informacion_Usar_Tarjeta_Verificacion = 28
    Checador_Informacion_Velocidad_Red = 29
    Checador_Informacion_Debera_Registrar_Tarjeta = 30
    Checador_Informacion_Tiempo_Retencion_Estatus_Temporal = 31
    Checador_Informacion_Tiempo_Retencion = 32
    Checador_Informacion_Tiempo_Retencion_Menu = 33
    Checador_Informacion_Formato_Fecha = 34
'    If I = 35 Then
'        If dwValue = 1 Then
'            ls2.AddItem "Whether is 1: 1 match or not?:YES"
'            ls2.Refresh
'        Else
'            ls2.AddItem "Whether is 1: 1 match or not?:NO"
'            ls2.Refresh
'        End If
'    End If
End Enum
'Enumeraciones para el estatus del dispositivo
Public Enum Checador_Estatus
    Checador_Estatus_No_Administradores = 1
    Checador_Estatus_No_Usuarios = 2
    Checador_Estatus_No_Plantilla = 3
    Checador_Estatus_No_Contrasenas = 4
    Checador_Estatus_No_Logs_Administradores = 5
    Checador_Estatus_No_Transacciones = 6
End Enum

Public Sub Inicializa()
    Call Conectar_Ayudante.Llena_Combo_Item("Equipo_ID, (CAST(No_Equipo as varchar) +' '+ Descripcion) as Equipo", "Cat_Equipos_Identificadores", Cmb_Dipositivos, 0, "No_Equipo")
    If Cmb_Dipositivos.ListCount > 0 Then
        Call Conectar_Ayudante.Asigna_Item_Combo(Equipo_ID, Cmb_Dipositivos)
    End If
    Dtp_Fecha_Dispositivo.Value = Now
    Dtp_Fecha_Dispositivo1.Value = Now
End Sub

Private Sub Btn_Actualizar_Firmware_Click()
Dim Archivo_Actualizacion As String
If LblEstatus.Caption = "Conectado" Then
    Cmd_Dispositivos.CancelError = True
    Cmd_Dispositivos.DialogTitle = "Seleccione el archivo de actualizacion"
    Cmd_Dispositivos.Flags = cdlOFNHideReadOnly
    Cmd_Dispositivos.Filter = "Actualizacion Firmware(*.cfg)|*.cfg"
    Cmd_Dispositivos.FilterIndex = 2
    'Cmd_Exportar.FileName = "Incidenias" & Format(Now, "ddMMyyHHmmss") & ".DAT"
    Cmd_Dispositivos.ShowOpen
    Archivo_Actualizacion = Cmd_Dispositivos.FileName
    'valida que el archivo exista
    If Len(Dir$(Archivo_Actualizacion)) > 0 Then
        If MsgBox("El archivo no existe ", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        Else
            If Checador_Actualizar_Firmware(Archivo_Actualizacion) Then
                MsgBox "Se ha actualizado la version del Firmware correctamente", vbInformation + vbOKOnly, Me.Caption
            Else
                MsgBox "No se pudo actualizar la version del Firmware", vbInformation + vbOKOnly, Me.Caption
            End If
        End If
    End If
Else
    MsgBox "El dispositivo no esta conectado favor de verificar", vbInformation + vbOKOnly, Me.Caption
End If
End Sub

Private Sub Btn_Apagar_Equipo_Click()
On Error GoTo HANDLER
    If LblEstatus.Caption = "Conectado" Then
        If MsgBox("Va a reiniciar el equipo, desea continuar", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            If Checador.PowerOffDevice(Maquina) Then
                MsgBox "El equipo se reinicio satisfactoriamente", vbInformation + vbOKOnly, Me.Caption
            Else
                MsgBox "El equipo no se pudo reiniciar", vbInformation + vbOKOnly, Me.Caption
            End If
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
HANDLER:
    MsgBox Er.Description
End Sub

Private Sub Btn_Cerrar_Conexion_Click()
    If LblEstatus.Caption = "Conectado" Then
        Checador_DesConectar_Checador_Net
        LblEstatus.Caption = "Desconectado"
    End If
End Sub

Private Sub Btn_Colocar_Fecha_Click()
Dim Anio As Double
Dim Mes As Double
Dim Dia As Double
Dim hora As Double
Dim minuto As Double
Dim segundo As Double

    If LblEstatus.Caption = "Conectado" Then
        Dtp_Fecha_Dispositivo1.Value = Now
        Anio = CLng(Year(Dtp_Fecha_Dispositivo1.Value))
        Mes = CLng(Month(Dtp_Fecha_Dispositivo1.Value))
        Dia = CLng(Day(Dtp_Fecha_Dispositivo1.Value))
        hora = CLng(Hour(Dtp_Fecha_Dispositivo1.Value))
        minuto = CLng(Minute(Dtp_Fecha_Dispositivo1.Value))
        segundo = CLng(Second(Dtp_Fecha_Dispositivo1.Value))
        Call Checador_Actualizar_Fecha(CDbl(Maquina), Anio, Mes, Dia, hora, minuto, segundo)
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
End Sub

Private Sub Btn_Configuracion_Avanzada_Click()

End Sub

Private Sub Btn_Generar_Respaldo_Click()
Dim error As Integer

'Selecciona la ruta del respaldo
Dim Ruta_Exportacion As String
Dim Nombre_Archivo As String
On Error GoTo HANDLER
    Cmd_Dispositivos.CancelError = True
    Cmd_Dispositivos.DialogTitle = "Seleccione el directorio"
    Cmd_Dispositivos.Flags = cdlOFNHideReadOnly
    Cmd_Dispositivos.Filter = "Archivos de Datos(*.dat)|*.dat"
    Cmd_Dispositivos.FilterIndex = 2
    Cmd_Dispositivos.FileName = "Respaldo.dat"
    Cmd_Dispositivos.ShowSave
    Ruta_Exportacion = Cmd_Dispositivos.FileName
    Nombre_Archivo = Cmd_Dispositivos.FileTitle
    If Cmd_Dispositivos.FileName <> "" And Nombre_Archivo <> "" Then
        If Checador.BackupData(Ruta_Exportacion) Then
            MsgBox "Información guardada correctamente.", vbInformation + vbOKOnly, Me.Caption
        Else
            MsgBox "No se pudo realizar el respaldo de la información", vbInformation + vbOKOnly, Me.Caption
        End If
        
    End If
  
  Exit Sub

HANDLER:
    Exit Sub
End Sub

Private Sub Btn_Excel_Click()
Dim Columna As Integer     'Indica que columna del grid es la que se esta consultando
Dim Fila As Integer        'Indica que fila del grid es la que se esta consultando
Dim Ruta_Archivo As String 'Obtiene la rura del archivo en donde se desea guardar
Dim CABECERA As String     'Indica el texto a exportar

On Error GoTo ErrHandler
    'Set CancelError is True
    MDIFrm_Apl_Principal.CommonDialog1.CancelError = True
    'Set flags
    MDIFrm_Apl_Principal.CommonDialog1.Flags = cdlOFNHideReadOnly
    'Set filters
    MDIFrm_Apl_Principal.CommonDialog1.Filter = "Archivos de Excel |*.XLS|"
    'Specify default filter
    MDIFrm_Apl_Principal.CommonDialog1.FilterIndex = 2
    'Display the Open dialog box
    MDIFrm_Apl_Principal.CommonDialog1.ShowSave
    'Display name of selected file
    Ruta_Archivo = MDIFrm_Apl_Principal.CommonDialog1.FileName
    Open Ruta_Archivo For Output As #1
        For Fila = 0 To Grid_Usuarios.Rows - 1
            For Columna = 0 To Grid_Usuarios.Cols - 1
                CABECERA = CABECERA & Chr(9) & Grid_Usuarios.TextMatrix(Fila, Columna)
            Next Columna
            Print #1, CABECERA
            CABECERA = ""
        Next Fila
    Close #1
    MsgBox "Reporte Exportado a Excel", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Ha ocurrido un error en la exportación, intentelo nuevamente", vbExclamation
    Exit Sub
End Sub

Private Sub Btn_Limpiar_Información_Click()
On Error GoTo HANDLER
    If LblEstatus.Caption = "Conectado" Then
        If MsgBox("Va a limpiar el log de acceso, huellas e información de usuario." _
            & Chr(13) & "(El dispositivo será reseteado y el proceso ya no podrá ser revertido)" _
            & Chr(13) & "¿Esta seguro de continuar de todas formas?", vbCritical + vbYesNo) = vbYes Then
            If Checador.ClearKeeperData(Maquina) Then
                MsgBox "Se realizó la limpieza de los registros", vbInformation + vbOKOnly, Me.Caption
            Else
                MsgBox "No se pudo realizar la limpieza de los registros", vbInformation + vbOKOnly, Me.Caption
            End If
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
Exit Sub
HANDLER:
    MsgBox Er.Description
End Sub

Private Sub Btn_Limpiar_Lista_Click()
    Grid_Usuarios.Rows = 0
End Sub

Private Sub Btn_LimpiarGLog_Click()
On Error GoTo HANDLER
    If LblEstatus.Caption = "Conectado" Then
        If MsgBox("Va a limpiar el log de los registros de acceso" & Chr(13) & "¿Esta seguro de continuar?", vbQuestion + vbYesNo) = vbYes Then
            If Checador.ClearGLog(Maquina) Then
                MsgBox "Se realizó la limpieza del log de registros", vbInformation + vbOKOnly, Me.Caption
            Else
                MsgBox "No se pudo realizar la limpieza del log de registros", vbInformation + vbOKOnly, Me.Caption
            End If
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
    Exit Sub
HANDLER:
    MsgBox Er.Description
End Sub

Private Sub Btn_Obtener_Fecha_Click()
    If LblEstatus.Caption = "Conectado" Then
        'If Checador_Conectar_Checador_Net(IP, Puerto) Then
        If Checador_Obtener_Fecha(Maquina, Dtp_Fecha_Dispositivo, Dtp_Fecha_Dispositivo1) Then
            Dtp_Fecha_Dispositivo1.Value = Now
        Else
            MsgBox "No se pudo obtener la fecha", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
End Sub

Private Sub Btn_Obtener_Usuarios_Click()
Dim res As Boolean
Dim errorCode As Long
Dim EnrollNumber As Long
Dim machinePrivilege As Long
Dim EnrollNum As Long
Dim Name As String
Dim Password As String
Dim Enabled As Boolean
Dim Cadena_Grid As String
    
    Grid_Usuarios.Rows = 0
    Grid_Usuarios.Cols = 5
    If LblEstatus.Caption = "Conectado" Then
        res = Checador.ReadAllUserID(CLng(Maquina))
        If res Then
            Grid_Usuarios.AddItem "ID" & Chr(9) & "Nombre" & Chr(9) & "Password" & Chr(9) & "Privilegios" & Chr(9) & "Habilitado"
            While Checador.GetAllUserInfo(Maquina, EnrollNum, Name, Password, machinePrivilege, Enabled)
                Cadena_Grid = EnrollNum & Chr(9)
                If Name = "" Then
                    Cadena_Grid = Cadena_Grid & "" & Chr(9)
                Else
                    Cadena_Grid = Cadena_Grid & Name & Chr(9)
                End If
                If Password = "" Then
                    Cadena_Grid = Cadena_Grid & "" & Chr(9)
                Else
                    Cadena_Grid = Cadena_Grid & Password & Chr(9)
                End If
                Select Case (machinePrivilege)
                    Case 0:
                        Cadena_Grid = Cadena_Grid & "Usuario General" & Chr(9)
                    Case 1:
                        Cadena_Grid = Cadena_Grid & "Admin Nivel 1" & Chr(9)
                    Case 2:
                        Cadena_Grid = Cadena_Grid & "Admin Nivel 2" & Chr(9)
                    Case 3:
                        Cadena_Grid = Cadena_Grid & "Admin Nivel 3" & Chr(9)
                    Case Else:
                        Cadena_Grid = Cadena_Grid & "Desconocido" & Chr(9)
                End Select
                If (Enabled = 1) Then
                    Cadena_Grid = Cadena_Grid & "S"
                Else
                    Cadena_Grid = Cadena_Grid & "N"
                End If
                Grid_Usuarios.AddItem Cadena_Grid
            Wend
            If Grid_Usuarios.Rows > 0 Then
                Grid_Usuarios.FixedRows = 1
            Else
                Grid_Usuarios.Rows = 0
            End If
        Else
            MsgBox "No se pudieron extraer los usuarios", vbInformation + vbOKOnly, Me.Caption
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
    Grid_Usuarios.ColWidth(0) = 1000
    Grid_Usuarios.ColWidth(1) = 3000
    Grid_Usuarios.ColWidth(2) = 1000
    Grid_Usuarios.ColWidth(3) = 2000
    Grid_Usuarios.ColWidth(4) = 1000
    Grid_Usuarios.Col = 0
    Grid_Usuarios.Sort = flexSortGenericAscending
End Sub

Private Sub Btn_Reinicar_Click()
On Error GoTo HANDLER
    If LblEstatus.Caption = "Conectado" Then
        If MsgBox("Va a reiniciar el equipo, desea continuar", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            If Checador.RestartDevice(Maquina) Then
                MsgBox "El equipo se reinicio satisfactoriamente", vbInformation + vbOKOnly, Me.Caption
            Else
                MsgBox "El equipo no se pudo reiniciar", vbInformation + vbOKOnly, Me.Caption
            End If
        End If
    Else
        MsgBox "El dispositivo no esta conectado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
    End If
HANDLER:
    MsgBox Er.Description
End Sub

Private Sub Cmb_Dipositivos_Click()
    If Cmb_Dipositivos.ListIndex > -1 Then
        Checador_DesConectar_Checador_Net
        'Llena informacion del dispositivo
        Call Llena_Informacion_Dispositivo(Format(Cmb_Dipositivos.ItemData(Cmb_Dipositivos.ListIndex), "00000"))
    End If
End Sub

'*******************************************************************************
'NOMBRE_FUNCION: Llena_Informacion_Dispositivo
'DESCRIPCION: Agrega la informacion general del dispositivo en las cajas de texto
'PARAMETROS : Checador_ID- Es el equipo que se va a conectar
'CREO       : Sergio Ulises Durán Hernández
'FECHA_CREO :
'MODIFICO   :
'FECHA_MODIFICO:
'CAUSA_MODIFICO:
'*******************************************************************************
Private Sub Llena_Informacion_Dispositivo(Checador_ID)
Dim Rs_Consulta_Informacion_Dispositivo As rdoResultset     'Informacion dek dispositivo

On Error GoTo HANDLER
    Me.MousePointer = 11
    'Consulta la informacion para conectarse
    Mi_SQL = "SELECT Direccion_IP,Puerto_IP,No_Equipo"
    Mi_SQL = Mi_SQL & " FROM Cat_Equipos_Identificadores"
    Mi_SQL = Mi_SQL & " WHERE Equipo_ID='" & Checador_ID & "'"
    Set Rs_Consulta_Informacion_Dispositivo = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
    If Not Rs_Consulta_Informacion_Dispositivo.EOF Then
        IP = Rs_Consulta_Informacion_Dispositivo.rdoColumns("Direccion_IP")
        Puerto = Rs_Consulta_Informacion_Dispositivo.rdoColumns("Puerto_IP")
        Maquina = Rs_Consulta_Informacion_Dispositivo.rdoColumns("No_Equipo")
    Else
        MsgBox "No hay informacion para el dispositivo seleccionado, favor de verificar", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    Set Rs_Consulta_Informacion_Dispositivo = Nothing
    'Realiza la conexion al dispositivo
    If Checador_Conectar_Checador_Net(IP, Puerto) Then
        'Realiza la obtencion de informacion
        Txt_Direccion_IP.Text = Checador_Obtener_IP(Maquina)
        Txt_Numero_Serie.Text = Checador_Obtener_No_Serie(Maquina)
        Txt_Firmware.Text = Checador_Obtener_Firmware(Maquina)
        Txt_Mac.Text = Checador_Obtener_MAC(Maquina)
        Txt_Formato_Fecha.Text = Checador_Formato_Fecha(Checador_Obtener_Informacion(Maquina, Checador_Informacion.Checador_Informacion_Formato_Fecha))
        Txt_No_Administradores.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Administradores)
        Txt_No_Usuarios.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Usuarios)
        Txt_No_Plantilla.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Plantilla)
        Txt_No_Contraseñas.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Contrasenas)
        Txt_Logs_Administrador.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Logs_Administradores)
        Txt_Transacciones.Text = Checador_Obtener_Estatus(Maquina, Checador_Estatus.Checador_Estatus_No_Transacciones)
        LblEstatus.Caption = "Conectado"
    Else
        MsgBox "No se puede conectar al dispositivo", vbInformation + vbOKOnly, Me.Caption
        LblEstatus.Caption = "Desconectado"
    End If
    Me.MousePointer = 0
Exit Sub
HANDLER:
    Me.MousePointer = 0
    MsgBox Err.Description, vbInformation + vbOKOnly, Me.Caption
End Sub

'*****************************Funciones para el checador***********************
Public Function Checador_Conectar_Checador_Net(IP_Checador As String, Puerto_Checador As Long) As Boolean
    Checador_Conectar_Checador_Net = Checador.Connect_Net(IP_Checador, Puerto_Checador)
End Function

Public Sub Checador_DesConectar_Checador_Net()
    Checador.Disconnect
End Sub

Public Function Checador_Agregar_Usuario(No_Maquina As Long, No_Usuario As Long, Nombre_Usuario As String, _
        Contrasena_Usuario As String, Privilegios As Long, Habilitado As Boolean) As Boolean
    Checador_Agregar_Usuario = Checador.SetUserInfo(No_Maquina, No_Usuario, Nombre_Usuario, Contrasena_Usuario, Privilegios, Habilitado)
End Function
Public Function Checador_Obtener_IP(No_Equipo As Long) As String
    If Checador.GetDeviceIP(No_Equipo, Checador_Obtener_IP) Then
        Exit Function
    Else
        Checador_Obtener_IP = "No se pudo obtener la direccion IP"
    End If
End Function
Public Function Checador_Obtener_MAC(No_Equipo As Long) As String
    If Checador.GetDeviceMAC(No_Equipo, Checador_Obtener_MAC) Then
        Exit Function
    Else
        Checador_Obtener_MAC = "No se pudo obtener la direccion MAC"
    End If
End Function

Public Function Checador_Obtener_Informacion(No_Equipo As Long, Opcion_Consultar As Long) As Long
Dim Informacion As Boolean
    Informacion = Checador.GetDeviceInfo(No_Equipo, Opcion_Consultar, Checador_Obtener_Informacion)
End Function

Public Function Checador_Obtener_Estatus(No_Equipo As Long, Opcion_Consultar As Long) As Long
Dim Informacion As Boolean
    Informacion = Checador.GetDeviceStatus(No_Equipo, Opcion_Consultar, Checador_Obtener_Estatus)
End Function


Public Function Checador_Obtener_Firmware(No_Equipo As Long) As String
    If Checador.GetFirmwareVersion(No_Equipo, Checador_Obtener_Firmware) Then
        Exit Function
    Else
        Checador_Obtener_Firmware = "No se pudo obtener el Firmware"
    End If
End Function
Public Function Checador_Obtener_No_Serie(No_Equipo As Long) As String
    If Checador.GetSerialNumber(No_Equipo, Checador_Obtener_No_Serie) Then
        Exit Function
    Else
        Checador_Obtener_No_Serie = "No se pudo obtener el No de Serie"
    End If
End Function

Public Function Checador_Formato_Fecha(No_Formato As Long) As String
    Select Case No_Formato
        Case 0: Checador_Formato_Fecha = "YY-MM-DD"
        Case 1: Checador_Formato_Fecha = "YY/MM/DD"
        Case 2: Checador_Formato_Fecha = "YY.MM.DD"
        Case 3: Checador_Formato_Fecha = "MM-DD-YY"
        Case 4: Checador_Formato_Fecha = "MM/DD/YY"
        Case 5: Checador_Formato_Fecha = "MM.DD.YY"
        Case 6: Checador_Formato_Fecha = "DD-MM-YY"
        Case 7: Checador_Formato_Fecha = "DD/MM/YY"
        Case 8: Checador_Formato_Fecha = "DD.MM.YY"
        Case 9: Checador_Formato_Fecha = "YYYYMMDD"
    End Select
End Function

Public Function Checador_Actualizar_Firmware(Ruta_Archivo As String) As Boolean
    Checador_Actualizar_Firmware = Checador.UpdateFirmware(Ruta_Archivo)
End Function

Public Function Checador_Obtener_Fecha(No_Maquina As Long, ByRef Dtp As DTPicker, ByRef Dtp1 As DTPicker) As Boolean
Dim Ano As Long
Dim Mes As Long
Dim Dia As Long
Dim Horas As Long
Dim Minutos As Long
Dim Segundos As Long
Dim Fecha As Date

    If Checador.GetDeviceTime(No_Maquina, Ano, Mes, Dia, Horas, Minutos, Segundos) Then
        Fecha = Format(CStr(Mes) & "/" & CStr(Dia) & "/" & CStr(Ano) & " " & CStr(Horas) & ":" & CStr(Minutos) & ":" & CStr(Segundos), "MM/dd/yyyy HH:mm:ss")
        Dtp.Value = Fecha
        Dtp1.Value = Fecha
        Checador_Obtener_Fecha = True
    Else
        Checador_Obtener_Fecha = False
    End If
End Function

Public Function Checador_Actualizar_Fecha(Maquina As Double, iYear As Double, iMonth As Double, iDay As Double, iHour As Double, iMinute As Double, iSecond As Double)
Dim res As Boolean
    res = Checador.SetDeviceTime2(Maquina, iYear, iMonth, iDay, iHour, iMinute, iSecond)
    If res Then
        Checador.ClearLCD
        Checador.EnableClock True
        Checador.EnableDevice Maquina, True
        MsgBox "Fecha y Hora actualizada"
    Else
        Checador.ClearLCD
        Checador.EnableClock True
        Checador.EnableDevice Maquina, True
        MsgBox "No actualizado"
    End If
End Function

