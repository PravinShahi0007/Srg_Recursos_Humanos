Attribute VB_Name = "Util"
'-------------------------------------------------------------------------------
'GrFinger Sample
'(c) 2005 Griaule Tecnologia Ltda.
'http://www.griaule.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Tecnologia makes no representations concerning either the merchantability
'of this software or the suitability of this sample for any particular purpose.
'
'THIS SAMPLE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR
'IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
'IN NO EVENT SHALL GRIAULE BE LIABLE FOR ANY DIRECT, INDIRECT,
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'You can download the free version of GrFinger directly from Griaule website.
'
'These notices must be retained in any copies of any part of this
'documentation and/or sample.
'
'-------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Support and fingerprint management routines
' -----------------------------------------------------------------------------------

Option Explicit
Option Base 0

' Raw image data type.
Public Type rawImage
    ' Image data.
    img As Variant
    ' Image width.
    width As Long
    ' Image height.
    height As Long
    ' Image resolution.
    res As Long
End Type

' Template data Type
Public Type TTemplate
    ' Template data
    tpt() As Byte
    ' Template size
    Size As Long
End Type

' Some constants to make our code cleaner
Public Const ERR_CANT_OPEN_BD = -999
Public Const ERR_INVALID_ID = -998
Public Const ERR_INVALID_TEMPLATE = -997

' The last acquired image.
Public raw As rawImage
' The template extracted from last acquired image.
Public Template As TTemplate
' Database class.
Public DB As DBClass

' -----------------------------------------------------------------------------------
' Support functions
' -----------------------------------------------------------------------------------
'Escribe un mensaje en el LOG
Public Sub writeLog(msg As String)
    Frm_Adm_Entrada_Comedor.lbLog.AddItem (msg)
    Frm_Adm_Entrada_Comedor.lbLog.ListIndex = Frm_Adm_Entrada_Comedor.lbLog.ListCount - 1
    Frm_Adm_Entrada_Comedor.lbLog.ListIndex = -1
End Sub

' Write and describe an error.
Public Sub writeError(errorCode As Long)
'    Select Case errorCode
'        Case GR_ERROR_INITIALIZE_FAIL
'            writeLog ("Fail to Initialize GrFingerX. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_NOT_INITIALIZED
'            writeLog ("The GrFingerX Library is not initialized. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_FAIL_LICENSE_READ
'            writeLog ("License not found. See manual for troubleshooting. (Error:" & Str(errorCode) & ")")
'            MsgBox ("License not found. See manual for troubleshooting.")
'        Case GR_ERROR_NO_VALID_LICENSE
'            writeLog ("The license is not valid. See manual for troubleshooting. (Error:" & Str(errorCode) & ")")
'            MsgBox ("The license is not valid. See manual for troubleshooting.")
'        Case GR_ERROR_NULL_ARGUMENT
'            writeLog ("The parameter have a null value. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_FAIL
'            writeLog ("Fail to create a GDI object. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_ALLOC
'            writeLog ("Fail to create a context. Cannot allocate memory. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_PARAMETERS
'            writeLog ("One or more parameters are out of bound. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_WRONG_USE
'            writeLog ("This function cannot be called at this time. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_EXTRACT
'            writeLog ("Template Extraction failed. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_SIZE_OFF_RANGE
'            writeLog ("Image is too larger or too short.  (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_RES_OFF_RANGE
'            writeLog ("Image have too low or too high resolution. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_CONTEXT_NOT_CREATED
'            writeLog ("The Context could not be created. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_INVALID_CONTEXT
'            writeLog ("The Context does not exist. (Error:" & Str(errorCode) & ")")
'
'        'Capture error codes
'        Case GR_ERROR_CONNECT_SENSOR
'            writeLog ("Error while connection to sensor. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_CAPTURING
'            writeLog ("Error while capturing from sensor. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_CANCEL_CAPTURING
'            writeLog ("Error while stop capturing from sensor. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_INVALID_ID_SENSOR
'            writeLog ("The idSensor is invalid. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_SENSOR_NOT_CAPTURING
'            writeLog ("The sensor is not capturing. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_INVALID_EXT
'            writeLog ("The File have a unknown extension. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_INVALID_FILENAME
'            writeLog ("The filename is invalid. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_INVALID_FILETYPE
'            writeLog ("The file type is invalid. (Error:" & Str(errorCode) & ")")
'        Case GR_ERROR_SENSOR
'            writeLog ("The sensor raise an error. (Error:" & Str(errorCode) & ")")
'
'        'Our error codes
'        Case ERR_INVALID_TEMPLATE
'            writeLog ("Invalid Template. (Error:" & Str(errorCode) & ")")
'        Case ERR_INVALID_ID
'            writeLog ("Invalid ID. (Error:" & Str(errorCode) & ")")
'        Case ERR_CANT_OPEN_BD
'            writeLog ("Unable to connect to DataBase. (Error:" & Str(errorCode) & ")")
'        Case Else
'            writeLog ("Error:" & Str(errorCode))
'    End Select
End Sub

' Check if we have a valid template
Public Function TemplateIsValid() As Boolean
    ' Check template size
    TemplateIsValid = (Template.Size > 0)
End Function

'-----------------------------------------------------------------------------------
'Main functions for fingerprint recognition management
'-----------------------------------------------------------------------------------
'Inicializa GrFinger ActiveX y las utilidades necesarias
Public Function InitializeGrFinger()
Dim err As Integer
    
    'Opening database
    Set DB = New DBClass
    If DB.OpenDB() = False Then
        InitializeGrFinger = ERR_CANT_OPEN_BD
        Exit Function
    End If
    'Crea un nuevo Template
    ReDim Template.tpt(GR_MAX_SIZE_TEMPLATE) As Byte
    'Crea a nueva "raw image"
    raw.width = 0
    raw.height = 0
    'Initializa librerias
    err = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.Initialize
    If err < 0 Then
        InitializeGrFinger = err
        Exit Function
    End If
    InitializeGrFinger = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.CapInitialize
End Function

'Cierra la librería y base de datos
Public Sub FinalizeGrFinger()
    'Librerias
    Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.CapFinalize
    Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.Finalize
    'Base de datos
    DB.closeDB
    Set DB = Nothing
End Sub

'Agrega el template a la base de datos
Public Function Enroll(No_Empleado As Long) As Integer
    'Checa si el template es válido
    If TemplateIsValid() Then
        'Agrga a la base de datos
        Enroll = DB.AddTemplate(Template.tpt, No_Empleado)
        Exit Function
    End If
    Enroll = -1
End Function

' Extract a fingerprint template from current image
Public Function ExtractTemplate() As Integer
    Dim ret As Integer
    
    ' Set initial buffer size and allocate it
    Template.Size = GR_MAX_SIZE_TEMPLATE
    ' reallocate template buffer
    ReDim Preserve Template.tpt(Template.Size)
    ret = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.Extract(raw.img, raw.width, raw.height, raw.res, Template.tpt, Template.Size, GR_DEFAULT_CONTEXT)
    ' if error, set template size to 0
    ' Result < 0 => extraction problem
    If ret < 0 Then Template.Size = 0
    ' Set real buffer size and free unecessary data
    ReDim Preserve Template.tpt(Template.Size)
    ExtractTemplate = ret
End Function

'Identifica la huella en la base de datos
Public Function Identify(ByRef score As Long) As Long
Dim ret As Long
Dim i As Integer
Dim rs As ADODB.Recordset
Dim tpt() As Byte

    
    'Inicia el proceso de identificación validando el template
    If Not TemplateIsValid() Then
        Identify = ERR_INVALID_TEMPLATE
        Exit Function
    End If
    ret = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.IdentifyPrepare(Template.tpt, GR_DEFAULT_CONTEXT)
    'Cacha el error
    If ret < 0 Then
        Identify = ret
        Exit Function
    End If
    'Obtiene de la base de datos los templates almacenados
    Set rs = DB.getTemplates
    'Recorre los registros en la base de datos
    Do Until rs.EOF
        'Obtiene el template actual
        tpt = rs("template")
        If Not (IsNull(tpt)) Then
            'Compara el template
            ret = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.Identify(tpt, score, GR_DEFAULT_CONTEXT)
            'Valida si coincide
            If ret = GR_MATCH Then
                Identify = rs("No_Empleado")     'Obtiene el número de empleado
                rs.Close
                Exit Function
            ElseIf ret < 0 Then
                Identify = ret
                Exit Function
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    'No lo encuentra y regresa el código "no match"
    Identify = GR_NOT_MATCH
End Function

'Verifica si existe la huella
Public Function Verify(ByVal ID As Long, ByRef score As Long) As Integer
Dim tpt() As Byte
    
    'Checa si es válido el template
    If Not TemplateIsValid() Then
        Verify = ERR_INVALID_TEMPLATE
        Exit Function
    End If
    'Obtiene el template de la base de datos
    tpt = DB.getTemplate(ID)
    'Checa si encontró el template
    If UBound(tpt) = 0 Then
        Verify = ERR_INVALID_ID
        Exit Function
    End If
    'Compara el template
    Verify = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.Verify(Template.tpt, tpt, score, GR_DEFAULT_CONTEXT)
End Function

' Show GrFinger version and type
Public Sub MessageVersion()
Dim majorVersion As Byte
Dim minorVersion As Byte
Dim ret As Long
Dim vStr As String
    
    majorVersion = 0
    minorVersion = 0
    vStr = ""
    ret = Frm_Adm_Entrada_Comedor.GrFingerXCtrl1.GetGrFingerVersion(majorVersion, minorVersion)
    If ret = GRFINGER_FULL Then vStr = "FULL"
    If ret = GRFINGER_LIGHT Then vStr = "LIGHT"
    'If ret = GRFINGER_FREE Then vStr = "FREE"
    Call MsgBox("The GrFinger DLL version is " & majorVersion & "." & minorVersion & "." & vbCrLf & "The license type is '" & vStr & "'.", , "GrFinger Version")
End Sub

