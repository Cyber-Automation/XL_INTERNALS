Attribute VB_Name = "VBD_Kit_Cryptography"
' | ========================================================================================================================= | '
' | __________  __________________________     ____________  _____________________  ___________________________________   __  | '
' | __  ____/ \/ /__  __ )__  ____/__  __ \    ___    |_  / / /__  __/_  __ \__   |/  /__    |__  __/___  _/_  __ \__  | / /  | '
' | _  /    __  /__  __  |_  __/  __  /_/ /    __  /| |  / / /__  /  _  / / /_  /|_/ /__  /| |_  /   __  / _  / / /_   |/ /   | '
' | / /___  _  / _  /_/ /_  /___  _  _, _/     _  ___ / /_/ / _  /   / /_/ /_  /  / / _  ___ |  /   __/ /  / /_/ /_  /|  /    | '
' | \____/  /_/  /_____/ /_____/  /_/ |_|______/_/  |_\____/  /_/    \____/ /_/  /_/  /_/  |_/_/    /___/  \____/ /_/ |_/     | '
' |                                     _/_____/                                                                              | '
' | ========================================================================================================================= | '

' +-[MODULE: VBD_Kit_Cryptography]--------------------------------------------+
' |                                                                           |
' | [ENGINEER]: Zeus_0x01                                                     |
' | [TELEGRAM]: @Zeus_0x01 (Public Name)                                      |
' | [DESCRIPTION]: Реализация алгоритмов шифрования через связку VBA + WinAPI |
' |                                                                           |
' +---------------------------------------------------------------------------+

' // <copyright file="VBD_Kit_Cryptography.bas" division="Cyber_Automation">
' // (C) Copyright 2024 Zeus_0x01 "{CC93FDF1-EC09-49AC-88B3-33F25F324851}"
' // </copyright>

'---------------------------------------------------------------------------'
' // Implemented Functionality (Реализованный функционал):
'    Forum -  https://www.script-coding.ru/threads/vbd_kit_cryptography.198/
'    GitHub - https://github.com/Cyber-Automation/XL_INTERNALS/
'---------------------------------------------------------------------------'
' // Release_Version (Версия компонента) - [01.01]
'---------------------------------------------------------------------------'

'================================================'
' // Конфигурация сборки компонента
'------------------------------------------------'
#Const BuildMode_WindowsNT_Only = True
#Const BuildMode_MacOS_Only = False
#Const BuildMode_x64Only = False
#Const BuildMode_x32Only = False
'------------------------------------------------'
#Const BuildMode_PrivateModule = False
'------------------------------------------------'
#Const BuildMode_Release = True
#Const BuildMode_UnSafeMode = False
#Const BuildMode_SilentMode = False
#Const BuildMode_ExtraValidation = True
#Const BuildMode_Enable_DoEvents = True
#Const BuildMode_WinAPI_SkipRuntimeChecks = False
'================================================'

'============================================================================'
' // Параметры компиляции и поведения кода в рамках данного компонента (.bas)
'----------------------------------------------------------------------------'
' {
    Option Explicit          ' // < Явное декларирование переменных >
    Option Compare Binary    ' // < Метод двоичного сравнения строк >
    Option Base 0            ' // < Вектора без индексного смещения >

    ' // < Инкапсуляция компонента >
    #If BuildMode_PrivateModule Then
        Option Private Module
    #End If
' }
'============================================================================'

'---------------------------------------------------------------------------------'
Private Const GUID_VBComponent As String = "{CC93FDF1-EC09-49AC-88B3-33F25F324851}"
'---------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------------------------------------------------------'
' // Определяем возможность компиляции компонента в соответствии с конфигурацией сборки

'````````````````````````````````````````````````````````````````````````````````````````````````````'
#Const Windows_NT = (Mac = 0&)  ' // Семейство ОС [Операционных Систем] (Windows_NT или Multics/Unix)
#Const x64_Soft = (Win64 <> 0&) ' // Разрядность MS Office (x64 и x32) или иного приложения с VBA
'````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
#If (BuildMode_x64Only And Not x64_Soft) And Not BuildMode_x32Only Then

    Public Sub Invalid_BuildConfiguration()
        Debug.Print vbNewLine & String$(91, "-")
        Debug.Print "Недопустимая конфигурация сборки проекта: " & vbNewLine
        Debug.Print "Невозможно скомпилировать компонент под MS Office x64, т.к. у Вас установлен MS Office x32!"
        Debug.Print "Измените константу ""BuildMode_x64Only = False"" для компиляции компонент под MS Office x32!"
        Debug.Print String$(91, "-") & vbNewLine
    End Sub

    Public Sub Auto_DetectConfiguration(): Call Setup_Configuration: End Sub

#ElseIf (BuildMode_x32Only And x64_Soft) And Not BuildMode_x64Only Then

    Public Sub Invalid_BuildConfiguration()
        Debug.Print vbNewLine & String$(91, "-")
        Debug.Print "Недопустимая конфигурация сборки проекта: " & vbNewLine
        Debug.Print "Невозможно скомпилировать компонент под MS Office x32, т.к. у Вас установлен MS Office x64!"
        Debug.Print "Измените константу ""BuildMode_x32Only = False"" для компиляции компонента под MS Office x64!"
        Debug.Print String$(91, "-") & vbNewLine
    End Sub

    Public Sub Auto_DetectConfiguration(): Call Setup_Configuration: End Sub

#ElseIf (BuildMode_WindowsNT_Only And Not Windows_NT) And Not BuildMode_MacOS_Only Then

    Public Sub Invalid_BuildConfiguration()
        Debug.Print vbNewLine & String$(116, "-")
        Debug.Print "Недопустимая конфигурация сборки проекта: " & vbNewLine
        Debug.Print "Невозможно скомпилировать компонент под Windows, т.к. у Вас установлена операционная систмема семейств Multics/Unix!"
        Debug.Print "Измените константу ""BuildMode_WindowsNT_Only = False"" для компиляции компонента под семейства ОС Multics/Unix!"
        Debug.Print String$(116, "-") & vbNewLine
    End Sub

    Public Sub Auto_DetectConfiguration(): Call Setup_Configuration: End Sub

#ElseIf (BuildMode_MacOS_Only And Windows_NT) And Not BuildMode_WindowsNT_Only Then

    Public Sub Invalid_BuildConfiguration()
        Debug.Print vbNewLine & String$(111, "-")
        Debug.Print "Недопустимая конфигурация сборки проекта: " & vbNewLine
        Debug.Print "Невозможно скомпилировать компонент под MacOS, т.к. у Вас установлена операционная система семейств Windows_NT!"
        Debug.Print "Измените константу ""BuildMode_MacOS_Only = False"" для компиляции компонента под семейства ОС Windows_NT!"
        Debug.Print String$(111, "-") & vbNewLine
    End Sub

    Public Sub Auto_DetectConfiguration(): Call Setup_Configuration: End Sub

#Else
'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------------------------'

'---------------------------------------'
Public Enum VBProject_Crypt_Algorithm
    Crypt_WinAPI_AES_CryptoNextGen = &H6
    Crypt_WinAPI_AES_CryptoAPI = &H7
    Crypt_VBACode_XOR = &H8
    Crypt_VBACode_AES_256 = &H9
End Enum
'---------------------------------------'

'-----------------------------------'
Public Enum VBProject_Type_Hardware
    Type_Win32_BaseBoard
    Type_Win32_Processor
    Type_Win32_VideoController
    Type_Win32_PhysicalMemory
    Type_Win32_Firmware
End Enum
'-----------------------------------'

'----------------------------------------------'
Public Enum MSOffice_Type_ContentFormat_Crypt
    MSDI_Undefined_Crypt = &HFFFFFFFC
    MSDI_Null_Crypt = &HFFFFFFFD
    MSDI_NullString_Crypt = &HFFFFFFFE
    MSDI_Nothing_Crypt = &HFFFFFFFF
    MSDI_Empty_Crypt = &H0
    MSDI_Bool_Crypt = &H1
    MSDI_Int8_Crypt = &H2
    MSDI_Int16_Crypt = &H4
    MSDI_Int32_Crypt = &H8
    MSDI_Int64_Crypt = &H10
    MSDI_FloatPoint32_Crypt = &H20
    MSDI_FloatPoint64Cur_Crypt = &H40
    MSDI_FloatPoint64Dbl_Crypt = &H56
    MSDI_FloatPoint112_Crypt = &H80
    MSDI_Date_Crypt = &H100
    MSDI_String_Crypt = &H200
    MSDI_Range_Crypt = &H400
    MSDI_Array_Crypt = &H800
    MSDI_Object_Crypt = &H1000
    MSDI_Collection_Crypt = &H2000
    MSDI_Dictionary_Crypt = &H4000
    MSDI_File_Crypt = &H8000
    MSDI_Folder_Crypt = &H10000
    MSDI_VBComponent_Crypt = &H20000
    MSDI_VBProject_Crypt = &H40000
    MSDI_UserDefType_Crypt = &H80000
    MSDI_NonExistent_Directory_Crypt = &H100000
    
    ' {
        MSDI_Procedure_Crypt = &H120000
        MSDI_AutoDetect_Crypt = &HFFFFFFAD
    ' }
End Enum
'----------------------------------------------'

'----------------------------------'
Public Enum RSA_PoolFunctions
    AESKey_ProtectWithPublicKey
    AESKey_UnprotectWithPrivateKey
    GenerateRandom_AESKeys
    GenerateRandom_RSAKeys
End Enum
'----------------------------------'

'----------------------------'
Public Enum modRSA_KeyLength
    AES_256 = 256
    AES_512 = 512
    RSA_2048 = 2048
    RSA_3072 = 3072
    RSA_4096 = 4096
    RSA_8192 = 8192
End Enum
'----------------------------'

'----------------------------'
Private Enum Format_SizeFile
    SizeFormat_Byte
    SizeFormat_KByte
    SizeFormat_MByte
    SizeFormat_GByte
End Enum
'----------------------------'

'-----------------------------------'
Private Enum MSOffice_Type_Encoding
    Type_ASCII
    Type_UTF_8
    Type_UTF_16
    Type_UTF_32
    Type_Unicode
End Enum
'-----------------------------------'

'--------------------------------------'
Private Enum FileSystem_Directory_Type
    DirType_File = &H0
    DirType_Folder = &H1
    DirType_Invalid = &HFFFFFFFF
    DirType_NotFound = &HFFFFFFFE
End Enum
'--------------------------------------'

'--------------------------------------------'
Private Enum WinAPI_CreateFile_DesiredAccess
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum
'--------------------------------------------'

'--------------------------------------------'
Private Enum WinAPI_CreateFile_ShareMode
    FILE_SHARE_DEFAULT = &H0
    FILE_SHARE_DELETE = &H4
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum
'--------------------------------------------'

'--------------------------------------------------'
Private Enum WinAPI_CreateFile_CreationDisposition
    CREATE_ALWAYS = &H2
    CREATE_NEW = &H1
    OPEN_ALWAYS = &H4
    OPEN_EXISTING = &H3
    TRUNCATE_EXISTING = &H5
End Enum
'--------------------------------------------------'

'--------------------------------------------------'
Private Enum WinAPI_CreateFile_FlagsAndAttributes
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_OFFLINE = &H1000
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    FILE_FLAG_NO_BUFFERING = &H20000000
    FILE_FLAG_OPEN_NO_RECALL = &H100000
    FILE_FLAG_OPEN_REPARSE_POINT = &H200000
    FILE_FLAG_OVERLAPPED = &H40000000
    FILE_FLAG_POSIX_SEMANTICS = &H1000000
    FILE_FLAG_RANDOM_ACCESS = &H10000000
    FILE_FLAG_SESSION_AWARE = &H800000
    FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    FILE_FLAG_WRITE_THROUGH = &H80000000
End Enum
'--------------------------------------------------'

'-----------------------------------'
Private Enum MSOffice_Type_Building
    Type_MSOffice_Build_Undefined
    Type_MSOffice_Build_NotSupport
    Type_MSOffice_2016_32Bit
    Type_MSOffice_2016_64Bit
    Type_MSOffice_2019_365_32Bit
    Type_MSOffice_2019_365_64Bit
End Enum
'-----------------------------------'

'--------------------------------------'
Private Enum MSOffice_Type_Application
    Type_MSOffice_App_Undefined
    Type_MSOffice_App_NotSupport
    Type_MSOffice_Excel
    Type_MSOffice_PowerPoint
    Type_MSOffice_Word
    Type_MSOffice_Access
    Type_MSOffice_Outlook
End Enum
'--------------------------------------'

'----------------------------------------'
Private Enum MSOffice_Type_Document_Group
    mso_Excel = &H2
    mso_PowerPoint = &H3
    mso_Word = &H4
    mso_Access = &H5
    mso_Outlook = &H6
End Enum
'----------------------------------------'

'-------------------------------------------------------'
Private Enum Security_ManagementCenter_Privileges
    Without_PrivilegesChanges = &HFFFFFFFF
    Enable_AllMacros_NotRecommended = &H1
    Disable_AllMacros = &H4
    Disable_AllMacros_WithNotification = &H2
    Disable_AllMacros_ExceptDigitallySignedMacros = &H3
End Enum
'-------------------------------------------------------'

'-------------------------------------------------------'
Private Enum Security_ManagementCenter_AccessObjectModel
    Without_AOMChanges = &HFFFFFFFF
    Access_Denied = &H0
    Access_Provided = &H1
End Enum
'-------------------------------------------------------'

'----------------------------------'
Private Enum Security_SettingMacro
    Privileges = &H0
    AccessObjectModel = &H1
End Enum
'----------------------------------'

'-------------------------------------'
Private Enum WinAPI_SystemErrors
    ERROR_SUCCESS = &H0
    ERROR_FILE_NOT_FOUND = &H2
    ERROR_PATH_NOT_FOUND = &H3
    ERROR_ACCESS_DENIED = &H5
    ERROR_INVALID_HANDLE = &H6
    ERROR_INVALID_PARAMETER = &H57
    ERROR_CALL_NOT_IMPLEMENTED = &H78
    ERROR_INSUFFICIENT_BUFFER = &H7A
    ERROR_MORE_DATA = &HEA
    ERROR_OPERATION_ABORTED = &H3E3
    ERROR_NO_MORE_ITEMS = &H103
End Enum
'-------------------------------------'

'-------------------------------------'
Private Enum Registry_HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
End Enum
'-------------------------------------'

'---------------------------------'
Private Enum Registry_AccessTypes
    KEY_ALL_ACCESS = &HF003F
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_READ = &H20019
    KEY_SET_VALUE = &H2
    KEY_WOW64_32KEY = &H200
    KEY_WOW64_64KEY = &H100
    KEY_WRITE = &H20006
End Enum
'---------------------------------'

'---------------------------------'
Private Enum Registry_ValueTypes
    REG_SZ = &H1
    REG_BINARY = &H3
    REG_DWORD = &H4
    REG_QWORD = &HB
    REG_MULTI_SZ = &H7
    REG_EXPAND_SZ = &H2
End Enum
'---------------------------------'

'--------------------------'
Private Type SHA1_FourBytes
    A As Byte
    B As Byte
    C As Byte
    D As Byte
End Type

Private Type SHA1_OneLong
    L As Long
End Type
'--------------------------'

'--------------------------'
Private Type Large_Integer
    Low_Part  As Long
    High_Part As Long
End Type
'--------------------------'

'---------------------------'
Private Type WinAPI_FileTime
    tm_LowDateTime  As Long
    tm_HighDateTime As Long
End Type
'---------------------------'

'------------------------------'
Private Type WinAPI_SystemTime
    tm_Year         As Integer
    tm_Month        As Integer
    tm_DayOfWeek    As Integer
    tm_Day          As Integer
    tm_Hour         As Integer
    tm_Minute       As Integer
    tm_Second       As Integer
    tm_Milliseconds As Integer
End Type
'------------------------------'

'------------------------------------------------'
Private Type Registry_SectionData
    Handle                As LongPtr
    Keys_Count            As Long
    Keys_MaxLen_Name      As Long
    Keys_MaxLen_Value     As Long
    Section_AccessRight   As Registry_AccessTypes
    Section_Count         As Long
    Section_HKey          As Registry_HKeys
    Section_LastWriteTime As WinAPI_SystemTime
    Section_MaxLen_Name   As Long
    Section_Path          As String
End Type
'------------------------------------------------'

'-------------------------------'
#Const Has_PtrSafe = (VBA7 <> 0)
'-------------------------------'

'--------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Kernel32.dll)

    Private Declare PtrSafe _
            Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
                     ByVal lpFileName As String, _
                     ByVal dwDesiredAccess As Long, _
                     ByVal dwShareMode As Long, _
                     ByVal lpSecurityAttributes As Long, _
                     ByVal dwCreationDisposition As Long, _
                     ByVal dwFlagsAndAttributes As Long, _
                     ByVal hTemplateFile As Long _
            ) As Long

    Private Declare PtrSafe _
            Function ReadFile Lib "kernel32.dll" ( _
                     ByVal hFile As Long, _
                     ByRef lpBuffer As Any, _
                     ByVal nNumberOfBytesToRead As Long, _
                     ByRef lpNumberOfBytesRead As Long, _
                     ByVal lpOverlapped As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CloseHandle Lib "kernel32.dll" ( _
                     ByVal hObject As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function GetFileSizeEx Lib "kernel32.dll" ( _
                     ByVal hFile As LongPtr, _
                     ByRef lpFileSize As Large_Integer _
            ) As Long

    Private Declare PtrSafe _
            Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
                     ByVal dwFlags As Long, _
                     ByVal lpSource As LongPtr, _
                     ByVal dwMessageId As Long, _
                     ByVal dwLanguageId As Long, _
                     ByVal lpBuffer As String, _
                     ByVal nSize As Long, _
                     ByVal Arguments As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function FileTimeToSystemTime Lib "kernel32.dll" ( _
                     ByRef lpFileTime As WinAPI_FileTime, _
                     ByRef lpSystemTime As WinAPI_SystemTime _
            ) As Long

    Private Declare PtrSafe _
            Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                ByRef Destination As Any, _
                ByRef Source As Any, _
                ByVal Length As LongPtr _
            )
            
    Private Declare PtrSafe _
            Sub RtlMoveMemory Lib "kernel32.dll" ( _
                ByRef Destination As Any, _
                ByRef Source As Any, _
                ByVal Length As LongPtr _
            )

#Else

    ' // Old MS Office or MacOS

#End If
'--------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Bcrypt.dll) | Cryptography API: Next Generation

    Private Declare PtrSafe _
            Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
                     ByRef phAlgorithm As LongPtr, _
                     ByVal pszAlgId As LongPtr, _
                     ByVal pszImplementation As LongPtr, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
                     ByVal hAlgorithm As LongPtr, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptGetProperty Lib "bcrypt.dll" ( _
                     ByVal hObject As LongPtr, _
                     ByVal pszProperty As LongPtr, _
                     ByRef pbOutput As Any, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dfFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptSetProperty Lib "bcrypt.dll" ( _
                     ByVal hObject As LongPtr, _
                     ByVal pszProperty As LongPtr, _
                     ByRef pbInput As Any, _
                     ByVal cbInput As Long, _
                     ByVal dfFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptCreateHash Lib "bcrypt.dll" ( _
                     ByVal hAlgorithm As LongPtr, _
                     ByRef phHash As LongPtr, _
                     ByRef pbHashObject As Any, _
                     ByVal cbHashObject As Long, _
                     ByVal pbSecret As LongPtr, _
                     ByVal cbSecret As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptHashData Lib "bcrypt.dll" ( _
                     ByVal hHash As LongPtr, _
                     ByRef pbInput As Any, _
                     ByVal cbInput As Long, _
                     Optional ByVal dwFlags As Long = 0 _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptFinishHash Lib "bcrypt.dll" ( _
                     ByVal hHash As LongPtr, _
                     ByRef pbOutput As Any, _
                     ByVal cbOutput As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptDestroyHash Lib "bcrypt.dll" ( _
                     ByVal hHash As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptGenRandom Lib "bcrypt.dll" ( _
                     ByVal hAlgorithm As LongPtr, _
                     ByRef pbBuffer As Any, _
                     ByVal cbBuffer As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptGenerateSymmetricKey Lib "bcrypt.dll" ( _
                     ByVal hAlgorithm As LongPtr, _
                     ByRef HKey As LongPtr, _
                     ByRef pbKeyObject As Any, _
                     ByVal cbKeyObject As Long, _
                     ByRef pbSecret As Any, _
                     ByVal cbSecret As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptEncrypt Lib "bcrypt.dll" ( _
                     ByVal HKey As LongPtr, _
                     ByRef pbInput As Any, _
                     ByVal cbInput As Long, _
                     ByRef pPaddingInfo As Any, _
                     ByRef pbIV As Any, _
                     ByVal cbIV As Long, _
                     ByRef pbOutput As Any, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptDecrypt Lib "bcrypt.dll" ( _
                     ByVal HKey As LongPtr, _
                     ByRef pbInput As Any, _
                     ByVal cbInput As Long, _
                     ByRef pPaddingInfo As Any, _
                     ByRef pbIV As Any, _
                     ByVal cbIV As Long, _
                     ByRef pbOutput As Any, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptDestroyKey Lib "bcrypt.dll" ( _
                     ByVal HKey As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function BCryptGenerateKeyPair Lib "bcrypt.dll" ( _
                    ByVal hAlgorithm As LongPtr, _
                    ByRef phKey As LongPtr, _
                    ByVal dwLength As Long, _
                    ByVal dwFlags As Long _
            ) As Long
    
    Private Declare PtrSafe _
            Function BCryptFinalizeKeyPair Lib "bcrypt.dll" ( _
                    ByVal HKey As LongPtr, _
                    ByVal dwFlags As Long _
            ) As Long
    
    Private Declare PtrSafe _
            Function BCryptExportKey Lib "bcrypt.dll" ( _
                    ByVal HKey As LongPtr, _
                    ByVal hExportKey As LongPtr, _
                    ByVal pszBlobType As LongPtr, _
                    ByVal pbOutput As LongPtr, _
                    ByVal cbOutput As Long, _
                    ByRef pcbResult As Long, _
                    ByVal dwFlags As Long _
            ) As Long
            
    Private Declare PtrSafe _
            Function BCryptImportKeyPair Lib "bcrypt.dll" ( _
                     ByVal hAlgorithm As LongPtr, _
                     ByVal hImportKey As LongPtr, _
                     ByVal pszBlobType As LongPtr, _
                     ByRef phKey As LongPtr, _
                     ByRef pbInput As Byte, _
                     ByVal cbInput As Long, _
                     ByVal dwFlags As Long _
            ) As Long
    
    Private Declare PtrSafe _
            Function BCryptExportKeyToBuffer Lib "bcrypt.dll" Alias "BCryptExportKey" ( _
                     ByVal HKey As LongPtr, _
                     ByVal hExportKey As LongPtr, _
                     ByVal pszBlobType As LongPtr, _
                     ByRef pbOutput As Byte, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dwFlags As Long _
            ) As Long
    
    Private Declare PtrSafe _
            Function BCryptEncryptToBuffer Lib "bcrypt.dll" Alias "BCryptEncrypt" ( _
                     ByVal HKey As LongPtr, _
                     ByRef pbInput As Byte, _
                     ByVal cbInput As Long, _
                     ByVal pPaddingInfo As LongPtr, _
                     ByVal pbIV As LongPtr, _
                     ByVal cbIV As Long, _
                     ByRef pbOutput As Byte, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dwFlags As Long _
            ) As Long
                    
    Private Declare PtrSafe _
            Function BCryptDecryptToBuffer Lib "bcrypt.dll" Alias "BCryptDecrypt" ( _
                     ByVal HKey As LongPtr, _
                     ByRef pbInput As Byte, _
                     ByVal cbInput As Long, _
                     ByVal pPaddingInfo As LongPtr, _
                     ByVal pbIV As LongPtr, _
                     ByVal cbIV As Long, _
                     ByRef pbOutput As Byte, _
                     ByVal cbOutput As Long, _
                     ByRef pcbResult As Long, _
                     ByVal dwFlags As Long _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'----------------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Advapi32.dll) | Crypto API

    Private Declare PtrSafe _
            Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
                     ByRef phProv As LongPtr, _
                     ByVal pszContainer As String, _
                     ByVal pszProvider As String, _
                     ByVal dwProvType As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptCreateHash Lib "advapi32.dll" ( _
                     ByVal hProv As LongPtr, _
                     ByVal Algid As Long, _
                     ByVal HKey As LongPtr, _
                     ByVal dwFlags As Long, _
                     ByRef phHash As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function CryptHashData Lib "advapi32.dll" ( _
                     ByVal hHash As LongPtr, _
                     ByRef pbData As Any, _
                     ByVal dwDataLen As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptDeriveKey Lib "advapi32.dll" ( _
                     ByVal hProv As LongPtr, _
                     ByVal Algid As Long, _
                     ByVal hBaseData As LongPtr, _
                     ByVal dwFlags As Long, _
                     ByRef phKey As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function CryptEncrypt Lib "advapi32.dll" ( _
                     ByVal HKey As LongPtr, _
                     ByVal hHash As LongPtr, _
                     ByVal Final As Long, _
                     ByVal dwFlags As Long, _
                     ByRef pbData As Any, _
                     ByRef pdwDataLen As Long, _
                     ByVal dwBufLen As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptDecrypt Lib "advapi32.dll" ( _
                     ByVal HKey As LongPtr, _
                     ByVal hHash As LongPtr, _
                     ByVal Final As Long, _
                     ByVal dwFlags As Long, _
                     ByRef pbData As Any, _
                     ByRef pdwDataLen As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptGetHashParam Lib "advapi32.dll" ( _
                     ByVal hHash As LongPtr, _
                     ByVal dwParam As Long, _
                     ByRef pbData As Any, _
                     ByRef pdwDataLen As Long, _
                     ByVal dwFlags As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptDestroyHash Lib "advapi32.dll" ( _
                     ByVal hHash As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function CryptDestroyKey Lib "advapi32.dll" ( _
                     ByVal HKey As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function CryptReleaseContext Lib "advapi32.dll" ( _
                     ByVal hProv As LongPtr, _
                     ByVal dwFlags As Long _
            ) As Long
            
    Private Declare PtrSafe _
            Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpSubKey As String, _
                     ByVal ulOptions As Long, _
                     ByVal samDesired As Long, _
                     ByRef phkResult As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpClass As String, _
                     ByRef lpcClass As Long, _
                     ByVal lpReserved As Long, _
                     ByRef lpcSubKeys As Long, _
                     ByRef lpcbMaxSubKeyLen As Long, _
                     ByRef lpcbMaxClassLen As Long, _
                     ByRef lpcValues As Long, _
                     ByRef lpcbMaxValueNameLen As Long, _
                     ByRef lpcbMaxValueLen As Long, _
                     ByRef lpcbSecurityDescriptor As Long, _
                     ByRef lpftLastWriteTime As WinAPI_FileTime _
            ) As Long

    Private Declare PtrSafe _
            Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpSubKey As String, _
                     ByVal Reserved As Long, _
                     ByVal lpType As Long, _
                     ByRef lpData As Any, _
                     ByVal lpcbData As Long _
            ) As Long

    Private Declare PtrSafe _
            Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpValueName As String, _
                     ByVal lpReserved As Long, _
                     ByRef lpType As Long, _
                     ByRef lpData As Any, _
                     ByRef lpcbData As Long _
            ) As Long

    Private Declare PtrSafe _
            Function RegCloseKey Lib "advapi32.dll" ( _
                     ByVal HKey As LongPtr _
            ) As Long
            
#Else

    ' // Old MS Office or MacOS

#End If
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Crypt32.dll) |  Crypto API

    Private Declare PtrSafe _
            Function CryptBinaryToString Lib "crypt32.dll" Alias "CryptBinaryToStringA" ( _
                     ByRef pbBinary As Any, _
                     ByVal cbBinary As Long, _
                     ByVal dwFlags As Long, _
                     ByVal pszString As String, _
                     ByRef pcchString As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CryptStringToBinary Lib "crypt32.dll" Alias "CryptStringToBinaryA" ( _
                     ByVal pszString As String, _
                     ByVal cchString As Long, _
                     ByVal dwFlags As Long, _
                     ByRef pbBinary As Any, _
                     ByRef pcbBinary As Long, _
                     Optional ByVal pdwSkip As Long = 0, _
                     Optional ByRef pdwFlags As Long = 0 _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Shell32.dll)

    Private Declare PtrSafe _
            Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" ( _
                     ByVal hWnd As LongPtr, _
                     ByVal pszPath As String, _
                     ByVal psa As Any _
            ) As Long
            
    Private Declare PtrSafe _
            Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                     ByVal hWnd As Long, _
                     ByVal lpOperation As String, _
                     ByVal lpFile As String, _
                     ByVal lpParameters As String, _
                     ByVal lpDirectory As String, _
                     ByVal nShowCmd As Long _
            ) As Long
            
#Else

    ' // Old MS Office or MacOS

#End If
'-------------------------------------------------------------------------------------------'

'------------------------------------------------------------------'
Private Glb_MSOffice_Type_Building    As MSOffice_Type_Building
Private Glb_MSOffice_Type_Application As MSOffice_Type_Application
'------------------------------------------------------------------'

'------------------------------------------------------'
Private Const CRYPT_STRING_BASE64 As Long = &H1
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000
'------------------------------------------------------'
Private Const STATUS_SUCCESS      As Long = 0
Private Const BCRYPT_PAD_PKCS1    As Long = &H2
'------------------------------------------------------'

'----------------------------------------------------------------------------------------------------'
Private Const MS_ENH_RSA_AES_PROV As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"
Private Const PROV_RSA_AES        As Long = 24&
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
'----------------------------------------------------------------------------------------------------'
Private Const ALG_CLASS_HASH  As Long = 32768
Private Const ALG_TYPE_ANY    As Long = 0&
Private Const ALG_SID_SHA_256 As Long = 12&
Private Const CALG_SHA_256    As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA_256)
'----------------------------------------------------------------------------------------------------'
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576&
Private Const ALG_TYPE_BLOCK     As Long = 1536&
Private Const ALG_SID_AES_256 As Long = 16&
Private Const CALG_AES_256 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_AES_256)
'----------------------------------------------------------------------------------------------------'
Private Const ALG_SID_SHA1  As Long = 4&
Private Const HP_HASHSIZE   As Long = 4&
Private Const HP_HASHVAL    As Long = 2&
Private Const PROV_RSA_FULL As Long = 1&
Private Const CALG_SHA1     As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA1)
'----------------------------------------------------------------------------------------------------'

'-------------------------------------------'
Private Const clOneMask   As Long = 16515072
Private Const clTwoMask   As Long = 258048
Private Const clThreeMask As Long = 4032&
Private Const clFourMask  As Long = 63&
'-------------------------------------------'
Private Const clHighMask  As Long = 16711680
Private Const clMidMask   As Long = 65280
Private Const clLowMask   As Long = 255&
'-------------------------------------------'
Private Const cl2Exp18    As Long = 262144
Private Const cl2Exp12    As Long = 4096&
Private Const cl2Exp6     As Long = 64&
Private Const cl2Exp8     As Long = 256&
Private Const cl2Exp16    As Long = 65536
'-------------------------------------------'

'---------------------------------------------------------------------------------------------------------------'
Private Const FILE_SYSTEM_SAVE_PATH   As String = _
                                      "\Cyber_Automation\[VBD_Kit] Компоненты разработчика\VBD_Kit_Cryptography\"
'---------------------------------------------------------------------------------------------------------------'
Private Const FILE_SYSTEM_LOCAL_APP_DATA              As String = "\AppData\Local"
Private Const FILE_SYSTEM_SAVE_FOLDER_ENCRYPTED_FILES As String = "\Encrypt_Files\"
Private Const FILE_SYSTEM_SAVE_FOLDER_DECRYPTED_FILES As String = "\Decrypt_Files\"
'---------------------------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
Private Const REGISTRY_SECTION_MSO As String = "SOFTWARE\Microsoft\Office\"

Private Const REGISTRY_SECTION_EXCEL_SECURITY      As String = "\Excel\Security\"
Private Const REGISTRY_SECTION_WORD_SECURITY       As String = "\Word\Security\"
Private Const REGISTRY_SECTION_ACCESS_SECURITY     As String = "\Access\Security\"
Private Const REGISTRY_SECTION_OUTLOOK_SECURITY    As String = "\Outlook\Security\"
Private Const REGISTRY_SECTION_POWERPOINT_SECURITY As String = "\PowerPoint\Security\"
'-------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------'
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000

Private Const LANG_NEUTRAL    As Long = &H0
Private Const SUBLANG_DEFAULT As Long = &H1
Private Const LANG_DEFAULT    As Long = (SUBLANG_DEFAULT * &H400 + LANG_NEUTRAL)
'-------------------------------------------------------------------------------'


'============================================================================================================'
Public Function Crypt_Protect_Data( _
                ByRef Source_Data As String, _
                ByRef Key_Crypt As String, _
                Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen, _
                Optional ByVal PIM As Long = 1, _
                Optional ByVal Use_Base64 As Boolean = True _
       ) As String
'------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-crypt_protect_data-vbd_kit_cryptography-bas.293/
'------------------------------------------------------------------------------------------------------------'

    '```````````````````````'
    Dim Data_Bytes() As Byte
    '```````````````````````'
    
    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    If Len(Source_Data) = 0 Then Exit Function
    If Len(Key_Crypt) = 0 Then Exit Function

    Crypt_Protect_Data = Source_Data
    Do Until PIM < 1: GoSub GS¦Crypt_Data: PIM = PIM - 1: Loop

    If Use_Base64 Then
        Crypt_Protect_Data = Base64_Encode_CryptoNG(Data_Bytes)
    Else
        Crypt_Protect_Data = Data_Bytes
    End If

    Exit Function
    '``````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````'
GS¦Crypt_Data:

    Select Case Crypt_Method
        Case Crypt_VBACode_XOR:              GoSub GS¦Crypt_VBACode_XOR
        Case Crypt_VBACode_AES_256:          GoSub GS¦Crypt_VBACode_AES256
        Case Crypt_WinAPI_AES_CryptoNextGen: GoSub GS¦Crypt_WinAPI_CryptoNG
        Case Crypt_WinAPI_AES_CryptoAPI:     GoSub GS¦Crypt_WinAPI_CryptoAPI
    End Select

    Return
'```````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_WinAPI_CryptoNG:

    Crypt_Protect_Data = AES_128_Encrypt_CryptoNG(Crypt_Protect_Data, Key_Crypt)
    Data_Bytes = Crypt_Protect_Data: PIM = 1

    Return
'```````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_WinAPI_CryptoAPI:

    Crypt_Protect_Data = StrConv(AES_256_Encrypt_CryptoAPI(Crypt_Protect_Data, Key_Crypt), vbUnicode)
    Data_Bytes = Crypt_Protect_Data: PIM = 1

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_VBACode_XOR:

    Crypt_Protect_Data = XOR_Crypt_VBACode(Crypt_Protect_Data, True, Key_Crypt)
    Data_Bytes = Crypt_Protect_Data: PIM = 1

    Return
'```````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_VBACode_AES256:

    Crypt_Protect_Data = AES_256_Crypt_VBACode(Crypt_Protect_Data, Key_Crypt, True)
    Data_Bytes = Crypt_Protect_Data

    Return
'``````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'============================================================================================================'
Public Function Crypt_Unprotect_Data( _
                ByRef Encrypted_Data As String, _
                ByRef Key_Crypt As String, _
                Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen, _
                Optional ByVal PIM As Long = 1, _
                Optional ByVal Use_Base64 As Boolean = True _
       ) As String
'------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-crypt_unprotect_data-vbd_kit_cryptography-bas.294/
'------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````'
    Dim Duplicate_EncryptedData As String
    '````````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If Len(Encrypted_Data) = 0 Then Exit Function
    If Len(Key_Crypt) = 0 Then Exit Function

    Crypt_Unprotect_Data = Encrypted_Data

    If Use_Base64 Then
        Crypt_Unprotect_Data = Base64_Decode_CryptoNG(Encrypted_Data)
    End If

    Do Until PIM < 1: GoSub GS¦Decrypt_Data: PIM = PIM - 1: Loop
    
    Exit Function
    '`````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_Data:

    Select Case Crypt_Method
        Case Crypt_VBACode_XOR:              GoSub GS¦Decrypt_VBACode_XOR
        Case Crypt_VBACode_AES_256:          GoSub GS¦Decrypt_VBACode_AES256
        Case Crypt_WinAPI_AES_CryptoNextGen: GoSub GS¦Decrypt_WinAPI_CryptoNG
        Case Crypt_WinAPI_AES_CryptoAPI:     GoSub GS¦Decrypt_WinAPI_CryptoAPI
    End Select

    Return
'`````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_WinAPI_CryptoNG:

    Crypt_Unprotect_Data = AES_128_Decrypt_CryptoNG(Crypt_Unprotect_Data, Key_Crypt): PIM = 1

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_WinAPI_CryptoAPI:
    
    Duplicate_EncryptedData = Crypt_Unprotect_Data
    Crypt_Unprotect_Data = AES_256_Decrypt_CryptoAPI(Crypt_Unprotect_Data, Key_Crypt): PIM = 1
    
    If Crypt_Unprotect_Data = Duplicate_EncryptedData Then Crypt_Unprotect_Data = vbNullString

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_VBACode_XOR:

    Crypt_Unprotect_Data = XOR_Crypt_VBACode(Crypt_Unprotect_Data, False, Key_Crypt)
    Crypt_Unprotect_Data = NullByte_Trim(Crypt_Unprotect_Data): PIM = 1

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_VBACode_AES256:

    Crypt_Unprotect_Data = AES_256_Crypt_VBACode(Crypt_Unprotect_Data, Key_Crypt, False)
    If PIM = 1 Then Crypt_Unprotect_Data = NullByte_Trim(Crypt_Unprotect_Data)

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'==============================================================================================================='
Public Function CryptEx_Protect_Data( _
                ByRef Source_Data As Variant, _
                ByVal Key_Crypt As String, _
                Optional ByVal Explicit_DataType As MSOffice_Type_ContentFormat_Crypt = MSDI_AutoDetect_Crypt, _
                Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen, _
                Optional ByVal PIM As Long = 1, _
                Optional ByVal Use_OS_Data As Boolean = False, _
                Optional ByVal Use_Hardware_Data As Boolean = False, _
                Optional ByVal Type_Hardware As VBProject_Type_Hardware = Type_Win32_BaseBoard, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal dw_Reserved_2 As Long = 0& _
       ) As Variant
'---------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-cryptex_protect_data-vbd_kit_cryptography-bas.291/
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - Include_SubFolders | Шифровать файлы в папках с учётом подпапок
' dw_Reserved_2 - ToDo_Parameter     | Зарезервировано без явного описания функционала
'---------------------------------------------------------------------------------------------------------------'
    
    '```````````````````````````````````````````````````````````'
    Dim Obj_VBProject  As Object, Obj_VBComponent As Object
    Dim Obj_CodeModule As Object, Module_Name     As String
    Dim VBProject      As Variant, VBComponent    As Variant
    '```````````````````````````````````````````````````````````'
    Dim Data_Type  As MSOffice_Type_ContentFormat_Crypt
    Dim Temp_Array()  As String, File_Path    As String
    Dim File_Details  As Variant, FS_SavePath As String
    '```````````````````````````````````````````````````````````'
    Dim Error_Message As String, Dimension_Count As Long
    Dim Data_Bytes() As Byte, Rng As Range, Base_64 As Boolean
    '```````````````````````````````````````````````````````````'
    Dim Procedure_Code      As String, Procedure_Name As String
    '```````````````````````````````````````````````````````````'
    Dim Crypt_ComponentName As String, Data_ToEncrypt As Variant
    Dim Current_ModuleName  As String, Count_VBComponent As Long
    '```````````````````````````````````````````````````````````'
    Dim Flag_SuccessfulEncryption As Boolean
    '```````````````````````````````````````````````````````````'
    Dim Start_Line As Long, End_Line As Long
    '```````````````````````````````````````````````````````````'
    Dim Inx_LB1    As Long, Inx_UB1  As Long
    Dim Inx_LB2    As Long, Inx_UB2  As Long
    Dim Inx_LB3    As Long, Inx_UB3  As Long
    '```````````````````````````````````````````````````````````'
    Dim i As Long, J As Long, N As Long, K As Long
    '```````````````````````````````````````````````````````````'

    '``````````````````````````````````'
    CryptEx_Protect_Data = vbNullString
    '``````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Cryptography
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'
    
    '```````````````````````````````````````'
    If Len(Key_Crypt) = 0 Then Exit Function
    '```````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````````'
    If Use_OS_Data Then
        Key_Crypt = Key_Crypt & "|" & Get_ID_OperationSystem(Crypt_Method)
    End If
    '`````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    If Use_Hardware_Data Then
        Key_Crypt = Key_Crypt & "|" & Get_ID_Hardware(Crypt_Method, Type_Hardware)
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    If Explicit_DataType = MSDI_AutoDetect_Crypt Then
        Data_Type = Get_SemanticDataType(Source_Data)
    Else
        Data_Type = Explicit_DataType
    End If

    Base_64 = True
    '````````````````````````````````````````````````'
    
    '``````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_Type
        Case MSDI_NonExistent_Directory_Crypt:     GoSub GS¦Exception_NonExistent_Directory
        Case MSDI_NullString_Crypt:                GoSub GS¦Exception_NullString
        Case MSDI_Bool_Crypt To MSDI_Date_Crypt, _
                                MSDI_String_Crypt: GoSub GS¦CryptEx_SimpleDataTypes
        Case MSDI_Range_Crypt:                     GoSub GS¦CryptEx_Range
        Case MSDI_Array_Crypt:                     GoSub GS¦CryptEx_Array
        Case MSDI_File_Crypt:                      GoSub GS¦CryptEx_File
        Case MSDI_Folder_Crypt:                    GoSub GS¦CryptEx_Folder
        Case MSDI_VBComponent_Crypt:               GoSub GS¦CryptEx_VBComponent
        Case MSDI_VBProject_Crypt:                 GoSub GS¦CryptEx_VBProject
        Case MSDI_Procedure_Crypt:                 GoSub GS¦CryptEx_Procedure
            
        Case Else
            Error_Message = "Тип данных не поддерживается: " & TypeName(Source_Data)
            Call Show_ErrorMessage_Immediate(Error_Message, "Несоответствие типов")
            CryptEx_Protect_Data = False

    End Select

    Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````'
GS¦CryptEx_SimpleDataTypes:
    
    Data_ToEncrypt = Source_Data: GoSub GS¦Crypt_Data

    Return
'````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````'
GS¦CryptEx_Range:
    
    For Each Rng In Source_Data
        Data_ToEncrypt = Rng.Formula: GoSub GS¦Crypt_Data
        If CryptEx_Protect_Data <> vbNullString Then
            Rng = CryptEx_Protect_Data
        End If
    Next Rng
    
    Set CryptEx_Protect_Data = Source_Data
    
    Return
'````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦CryptEx_Array:

    Dimension_Count = Get_ArrayDimension(Source_Data)
    Select Case Dimension_Count
        Case 0
            Error_Message = "Массив не инициализирован!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Нет данных для обработки")
            CryptEx_Protect_Data = False

        Case 1
            Inx_LB1 = LBound(Source_Data, 1)
            Inx_UB1 = UBound(Source_Data, 1)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1)

            For i = Inx_LB1 To Inx_UB1
                Data_ToEncrypt = Source_Data(i): GoSub GS¦Crypt_Data
                Temp_Array(i) = CryptEx_Protect_Data
            Next i

            CryptEx_Protect_Data = Temp_Array

        Case 2
            Inx_LB1 = LBound(Source_Data, 1): Inx_UB1 = UBound(Source_Data, 1)
            Inx_LB2 = LBound(Source_Data, 2): Inx_UB2 = UBound(Source_Data, 2)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1, Inx_LB2 To Inx_UB2)

            For J = Inx_LB2 To Inx_UB2
                For i = Inx_LB1 To Inx_UB1
                    Data_ToEncrypt = Source_Data(i, J): GoSub GS¦Crypt_Data
                    Temp_Array(i, J) = CryptEx_Protect_Data
                Next i
            Next J

            CryptEx_Protect_Data = Temp_Array

        Case 3
            Inx_LB1 = LBound(Source_Data, 1): Inx_UB1 = UBound(Source_Data, 1)
            Inx_LB2 = LBound(Source_Data, 2): Inx_UB2 = UBound(Source_Data, 2)
            Inx_LB3 = LBound(Source_Data, 3): Inx_UB3 = UBound(Source_Data, 3)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1, Inx_LB2 To Inx_UB2, Inx_LB3 To Inx_UB3)

            For K = Inx_LB3 To Inx_UB3
                For J = Inx_LB2 To Inx_UB2
                    For i = Inx_LB1 To Inx_UB1
                        Data_ToEncrypt = Source_Data(i, J, K): GoSub GS¦Crypt_Data
                        Temp_Array(i, J, K) = CryptEx_Protect_Data
                    Next i
                Next J
            Next K

            CryptEx_Protect_Data = Temp_Array

        Case Is > 3
            Error_Message = "Отсуствует поддержка массивов, больше 3-х размерностью!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Невозможно обработать массив")
            CryptEx_Protect_Data = False

        Case -1
            Error_Message = "Проблема идентификации массива! Проверьте, что Вы передаёте на обработку."
            Call Show_ErrorMessage_Immediate(Error_Message, "Невозможно обработать массив")
            CryptEx_Protect_Data = False

    End Select

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````'
GS¦CryptEx_File:

    GoSub GS¦Check_SavePath: File_Path = Source_Data
    Base_64 = False: GoSub GS¦Crypt_File
    
    If Not IsEmpty(CryptEx_Protect_Data) Then
        CryptEx_Protect_Data = True
    Else
        CryptEx_Protect_Data = False
    End If

    Return
'````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````'
GS¦CryptEx_Folder:

    GoSub GS¦Check_SavePath

    If Right(Source_Data, 1) <> "\" Then Source_Data = Source_Data & "\"

    File_Path = Dir(Source_Data): Base_64 = False
    Flag_SuccessfulEncryption = True

    Do While File_Path <> ""
        File_Path = Source_Data & File_Path
        GoSub GS¦Crypt_File: File_Path = Dir
        If IsEmpty(CryptEx_Protect_Data) Then: Flag_SuccessfulEncryption = False
    Loop
    
    If Flag_SuccessfulEncryption Then
        CryptEx_Protect_Data = True
    Else
        CryptEx_Protect_Data = False
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````'
GS¦CryptEx_VBProject:

    If Application.VBE.ActiveVBProject.Name = Source_Data Then
        Set Obj_VBProject = Application.VBE.ActiveVBProject
        Current_ModuleName = Application.VBE.ActiveCodePane.CodeModule.Name
    Else
        For Each VBProject In Application.VBE.VBProjects
            If VBProject.Name = Source_Data Then
                Set Obj_VBProject = VBProject: Exit For
            End If
        Next VBProject
    End If

    Count_VBComponent = Obj_VBProject.VBComponents.Count
    Crypt_ComponentName = Find_ModuleByGUID(GUID_VBComponent).Name

    For Each Obj_VBComponent In Obj_VBProject.VBComponents
        If Obj_VBComponent.Name <> Crypt_ComponentName Then
            If Len(Current_ModuleName) = 0 Then
                GoSub GS¦Crypt_Component
            Else
                If Current_ModuleName <> Obj_VBComponent.Name Then
                    GoSub GS¦Crypt_Component
                End If
            End If
        End If
    Next Obj_VBComponent

    CryptEx_Protect_Data = True
            
    Return
'```````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````'
GS¦CryptEx_VBComponent:

    For Each VBComponent In Application.VBE.ActiveVBProject.VBComponents
        If VBComponent.Name = Source_Data Then
            Set Obj_VBProject = Application.VBE.ActiveVBProject: Exit For
        End If
    Next

    If Obj_VBProject Is Nothing Then
        For Each VBProject In Application.VBE.VBProjects
            For Each VBComponent In VBProject.VBComponents
                If VBComponent.Name = Source_Data Then
                    Set Obj_VBProject = VBProject: Exit For
                End If
            Next VBComponent
            If Not Obj_VBProject Is Nothing Then Exit For
        Next VBProject
    End If

    Set Obj_VBComponent = Obj_VBProject.VBComponents(Source_Data)
    GoSub GS¦Crypt_Component: CryptEx_Protect_Data = True
    
    Return
'```````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦CryptEx_Procedure:
    
    If InStr(1, Source_Data, ".") = 0 Then
        Error_Message = _
        "Проверьте входные данные на соответствие шаблону - ""Module.Procedure""" & _
        "Текущий шаблон (Source_Data = " & Source_Data & ")"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка определения процедуры")
    End If
    
    Module_Name = Split(Source_Data, ".")(0): Procedure_Name = Split(Source_Data, ".")(1)
    If Not IsModule_Exists(Module_Name) Then
        Error_Message = "Модуль """ & Module_Name & """ отсутствует в текущем проекте VBA!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска модуля")
        CryptEx_Protect_Data = False: Exit Function
    End If
    
    If Not IsProcedure_Exists(Module_Name, Procedure_Name) Then
        Error_Message = "Процедура """ & Procedure_Name & """ отсутствует в модуле """ & _
                         Module_Name & """" & "! Возможно процедура имеет нестандартную сигнатуру!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска процедуры")
        CryptEx_Protect_Data = False: Exit Function
    End If
    
    Set Obj_VBProject = ThisWorkbook.VBProject
    Set Obj_VBComponent = Obj_VBProject.VBComponents(Module_Name)
    Set Obj_CodeModule = Obj_VBComponent.CodeModule
    
    Start_Line = Find_ProcedureStart(Obj_CodeModule, Procedure_Name)
    If Start_Line = 0 Then
        Error_Message = "Не удалось найти процедуру """ & Procedure_Name & """! " & _
                        "Возможно процедура имеет нестандартную сигнатура!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска процедуры")
        CryptEx_Protect_Data = False: Return
    End If
    
    End_Line = Find_ProcedureEnd(Obj_CodeModule, Start_Line)
    If End_Line = 0 Then
        Error_Message = "Не удалось найти конец процедуры """ & Procedure_Name & """! " & _
                        "Возможно процедура имеет нестандартную сигнатура завершения!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска процедуры")
        CryptEx_Protect_Data = False: Return
    End If
    
    Procedure_Code = Obj_CodeModule.Lines(Start_Line, End_Line - Start_Line + 1)
    Data_ToEncrypt = Procedure_Code: GoSub GS¦Crypt_Data
    
    Temp_Array = Split(CryptEx_Protect_Data, Chr(13) & Chr(10))
    For N = LBound(Temp_Array, 1) To UBound(Temp_Array, 1)
        Temp_Array(N) = "' // " & Temp_Array(N)
    Next N

    Data_ToEncrypt = Join(Temp_Array, Chr(13) & Chr(10))
    Data_ToEncrypt = "' // Procedure: " & Procedure_Name & " " & GUID_VBComponent & vbNewLine & _
                           Data_ToEncrypt & vbNewLine & "' // End " & GUID_VBComponent

    Obj_CodeModule.DeleteLines Start_Line, End_Line - Start_Line + 1
    Obj_CodeModule.InsertLines Start_Line, Data_ToEncrypt
    
    CryptEx_Protect_Data = True
    
    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_File:

    Data_ToEncrypt = Conversion_FileToBytes(File_Path)
    If IsEmpty(Data_ToEncrypt) Then
        CryptEx_Protect_Data = Empty
        Return
    End If
    
    GoSub GS¦Crypt_Data: Data_Bytes = CryptEx_Protect_Data

    Inx_LB1 = LBound(Data_Bytes, 1)
    Inx_UB1 = UBound(Data_Bytes, 1)

    ReDim Preserve Data_Bytes(Inx_LB1 To Inx_UB1 + 8)
    File_Details = Get_PathFileNameAndExtension(File_Path)

    For N = 1 To Len(File_Details(2))
        Data_Bytes(Inx_UB1 + N) = AscB(Mid$(File_Details(2), N, 1))
    Next N
    
    File_Details(0) = FS_SavePath
    File_Path = File_Details(0) & "\" & File_Details(1) & ".vbfc"

    If File_IsOpen(File_Path) Then
        Error_Message = "Копия файла в папке с обработанными файлами " & _
                        "на данным момент открыта и не может быть заменена!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу для замены")
        CryptEx_Protect_Data = Empty: Return
    End If
    
    CryptEx_Protect_Data = Conversion_BytesToFile(Data_Bytes, File_Path)

    Return
'`````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````'
GS¦Crypt_Component:

    Set Obj_CodeModule = Obj_VBComponent.CodeModule
    
    With Obj_CodeModule
        If .CountOfLines > 0 Then
            Data_ToEncrypt = .Lines(1, .CountOfLines)
            GoSub GS¦Crypt_Data: Data_ToEncrypt = CryptEx_Protect_Data

            Temp_Array = Split(Data_ToEncrypt, Chr(13) & Chr(10))
            For N = LBound(Temp_Array, 1) To UBound(Temp_Array, 1)
                Temp_Array(N) = "' // " & Temp_Array(N)
            Next N

            Data_ToEncrypt = Join(Temp_Array, Chr(13) & Chr(10))
        
            .DeleteLines 1, .CountOfLines
            .AddFromString Data_ToEncrypt
        End If
    End With

    Set Obj_CodeModule = Nothing

    Return
'`````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````'
GS¦Crypt_Data:

    CryptEx_Protect_Data = Crypt_Protect_Data( _
                           CStr(Data_ToEncrypt), Key_Crypt, Crypt_Method, PIM, Base_64)

    Return
'``````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Check_SavePath:
    
    FS_SavePath = FS_SavePath & Environ$("USERPROFILE")
    FS_SavePath = FS_SavePath & FILE_SYSTEM_LOCAL_APP_DATA
    FS_SavePath = FS_SavePath & FILE_SYSTEM_SAVE_PATH
    FS_SavePath = FS_SavePath & FILE_SYSTEM_SAVE_FOLDER_ENCRYPTED_FILES
    
    If Dir(FS_SavePath, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, FS_SavePath, ByVal 0&)
    End If

    Return
'```````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NonExistent_Directory:

    Error_Message = "Файл или папка не существует: " & Source_Data
    Call Show_ErrorMessage_Immediate(Error_Message, _
                                    "Ошибка поиска нужного файла/директории", , True)
    CryptEx_Protect_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NullString:

    Error_Message = "Нет данных для обработки: Len String = 0"
    Call Show_ErrorMessage_Immediate(Error_Message, "Передана пустая строка")
    CryptEx_Protect_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------------'
End Function
'==============================================================================================================='


'==================================================================================================================='
Public Function CryptEx_Unprotect_Data( _
                ByRef Encrypted_Data As Variant, _
                ByVal Key_Crypt As String, _
                Optional ByVal Explicit_DataType As MSOffice_Type_ContentFormat_Crypt = MSDI_AutoDetect_Crypt, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen, _
                Optional ByVal PIM As Long = 1, _
                Optional ByVal Use_OS_Data As Boolean = False, _
                Optional ByVal Use_Hardware_Data As Boolean = False, _
                Optional ByVal Type_Hardware As VBProject_Type_Hardware = Type_Win32_BaseBoard, _
                Optional ByVal dw_Reserved_2 As Long = 0& _
       ) As Variant
'-------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-cryptex_unprotect_data-vbd_kit_cryptography-bas.290/
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - ToDo_Parameter     | Зарезервировано без явного описания функционала
' dw_Reserved_2 - ToDo_Parameter     | Зарезервировано без явного описания функционала
'-------------------------------------------------------------------------------------------------------------------'
    
    '```````````````````````````````````````````````````````````'
    Dim Obj_VBProject  As Object, Obj_VBComponent As Object
    Dim Obj_CodeModule As Object, Module_Name     As String
    Dim VBProject      As Variant, VBComponent    As Variant
    '```````````````````````````````````````````````````````````'
    Dim Data_Type  As MSOffice_Type_ContentFormat_Crypt
    Dim Temp_Array()  As String, File_Path    As String
    Dim File_Details  As Variant, FS_SavePath As String
    '```````````````````````````````````````````````````````````'
    Dim Error_Message As String, Dimension_Count As Long
    Dim Data_Bytes() As Byte, Rng As Range, Base_64 As Boolean
    '```````````````````````````````````````````````````````````'
    Dim Procedure_Name As String, Encrypted_Lines  As Collection
    '```````````````````````````````````````````````````````````'
    Dim Crypt_ComponentName As String, Data_ToDecrypt As Variant
    Dim Current_ModuleName  As String, Count_VBComponent As Long
    '```````````````````````````````````````````````````````````'
    Dim Flag_SuccessfulDecryption As Boolean
    '```````````````````````````````````````````````````````````'
    Dim Start_Line As Long, End_Line As Long
    '```````````````````````````````````````````````````````````'
    Dim Inx_LB1    As Long, Inx_UB1  As Long
    Dim Inx_LB2    As Long, Inx_UB2  As Long
    Dim Inx_LB3    As Long, Inx_UB3  As Long
    '```````````````````````````````````````````````````````````'
    Dim i As Long, J As Long, N As Long, K As Long
    '```````````````````````````````````````````````````````````'
    
    '````````````````````````````````````'
    CryptEx_Unprotect_Data = vbNullString
    '````````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Cryptography
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'
    
    '```````````````````````````````````````'
    If Len(Key_Crypt) = 0 Then Exit Function
    '```````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````````'
    If Use_OS_Data Then
        Key_Crypt = Key_Crypt & "|" & Get_ID_OperationSystem(Crypt_Method)
    End If
    '`````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    If Use_Hardware_Data Then
        Key_Crypt = Key_Crypt & "|" & Get_ID_Hardware(Crypt_Method, Type_Hardware)
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    If Explicit_DataType = MSDI_AutoDetect_Crypt Then
        Data_Type = Get_SemanticDataType(Encrypted_Data)
    Else
        Data_Type = Explicit_DataType
    End If

    Base_64 = True
    '```````````````````````````````````````````````````'
    
    '``````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_Type
        Case MSDI_NonExistent_Directory_Crypt:     GoSub GS¦Exception_NonExistent_Directory
        Case MSDI_NullString_Crypt:                GoSub GS¦Exception_NullString
        Case MSDI_Bool_Crypt To MSDI_Date_Crypt, _
                                MSDI_String_Crypt: GoSub GS¦DeCryptEx_SimpleDataTypes
        Case MSDI_Range_Crypt:                     GoSub GS¦DeCryptEx_Range
        Case MSDI_Array_Crypt:                     GoSub GS¦DeCryptEx_Array
        Case MSDI_File_Crypt:                      GoSub GS¦DeCryptEx_File
        Case MSDI_Folder_Crypt:                    GoSub GS¦DeCryptEx_Folder
        Case MSDI_VBComponent_Crypt:               GoSub GS¦DeCryptEx_VBComponent
        Case MSDI_VBProject_Crypt:             GoSub GS¦DeCryptEx_VBProject
        Case MSDI_Procedure_Crypt:                 GoSub GS¦DeCryptEx_Procedure
        
        Case Else
            Error_Message = "Тип данных не поддерживается: " & TypeName(Encrypted_Data)
            Call Show_ErrorMessage_Immediate(Error_Message, "Несоответствие типов")
            CryptEx_Unprotect_Data = False

    End Select

    Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````'
    
'``````````````````````````````````````````````````````````'
GS¦DeCryptEx_SimpleDataTypes:
    
    Data_ToDecrypt = Encrypted_Data: GoSub GS¦Decrypt_Data
    
    Return
'``````````````````````````````````````````````````````````'
    
'``````````````````````````````````````````````````````````'
GS¦DeCryptEx_Range:

    For Each Rng In Encrypted_Data
        Data_ToDecrypt = Rng.Formula: GoSub GS¦Decrypt_Data
        If CryptEx_Unprotect_Data <> vbNullString Then
            Rng = CryptEx_Unprotect_Data
        End If
    Next Rng
    Set CryptEx_Unprotect_Data = Encrypted_Data

    Return
'``````````````````````````````````````````````````````````'
    
'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_Array:

    Dimension_Count = Get_ArrayDimension(Encrypted_Data)
    
    Select Case Dimension_Count
        Case 0
            Error_Message = "Массив не инициализирован!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Нет данных для обработки")
            CryptEx_Unprotect_Data = False

        Case 1
            Inx_LB1 = LBound(Encrypted_Data, 1)
            Inx_UB1 = UBound(Encrypted_Data, 1)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1)

            For i = Inx_LB1 To Inx_UB1
                Data_ToDecrypt = Encrypted_Data(i): GoSub GS¦Decrypt_Data
                Temp_Array(i) = CryptEx_Unprotect_Data
            Next i

            CryptEx_Unprotect_Data = Temp_Array

        Case 2
            Inx_LB1 = LBound(Encrypted_Data, 1): Inx_UB1 = UBound(Encrypted_Data, 1)
            Inx_LB2 = LBound(Encrypted_Data, 2): Inx_UB2 = UBound(Encrypted_Data, 2)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1, Inx_LB2 To Inx_UB2)

            For J = Inx_LB2 To Inx_UB2
                For i = Inx_LB1 To Inx_UB1
                    Data_ToDecrypt = Encrypted_Data(i, J): GoSub GS¦Decrypt_Data
                    Temp_Array(i, J) = CryptEx_Unprotect_Data
                Next i
            Next J

            CryptEx_Unprotect_Data = Temp_Array

        Case 3
            Inx_LB1 = LBound(Encrypted_Data, 1): Inx_UB1 = UBound(Encrypted_Data, 1)
            Inx_LB2 = LBound(Encrypted_Data, 2): Inx_UB2 = UBound(Encrypted_Data, 2)
            Inx_LB3 = LBound(Encrypted_Data, 3): Inx_UB3 = UBound(Encrypted_Data, 3)

            ReDim Temp_Array(Inx_LB1 To Inx_UB1, Inx_LB2 To Inx_UB2, Inx_LB3 To Inx_UB3)

            For K = Inx_LB3 To Inx_UB3
                For J = Inx_LB2 To Inx_UB2
                    For i = Inx_LB1 To Inx_UB1
                        Data_ToDecrypt = Encrypted_Data(i, J, K): GoSub GS¦Decrypt_Data
                        Temp_Array(i, J, K) = CryptEx_Unprotect_Data
                    Next i
                Next J
            Next K

            CryptEx_Unprotect_Data = Temp_Array

        Case Is > 3
            Error_Message = "Отсуствует поддержка массивов, больше 3-х размерностью!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Невозможно обработать массив")
            CryptEx_Unprotect_Data = False

        Case -1
            Error_Message = "Проблема идентификации массива! Проверьте, что Вы передаёте на обработку."
            Call Show_ErrorMessage_Immediate(Error_Message, "Невозможно обработать массив")
            CryptEx_Unprotect_Data = False

    End Select
            
    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'
    
'`````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_File:
    
    GoSub GS¦Check_SavePath
    File_Path = Encrypted_Data: Base_64 = False: GoSub GS¦Decrypt_File
    
    Return
'`````````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_Folder:
    
    GoSub GS¦Check_SavePath
    If Right(Encrypted_Data, 1) <> "\" Then Encrypted_Data = Encrypted_Data & "\"

    File_Path = Dir(Encrypted_Data & "*.vbfc"): Base_64 = False
    Flag_SuccessfulDecryption = True
     
    If Len(File_Path) = 0 Then
        Error_Message = "В выбранной папке нет файлов с расширением .vbfc (" & Encrypted_Data & ")"
        Call Show_ErrorMessage_Immediate(Error_Message, "Отсутствуют файлы для дешифрования")
        CryptEx_Unprotect_Data = False
    Else
        Do While File_Path <> ""
            File_Path = Encrypted_Data & File_Path
            GoSub GS¦Decrypt_File: File_Path = Dir
            If Not CryptEx_Unprotect_Data Then: Flag_SuccessfulDecryption = False
        Loop
        
        If Flag_SuccessfulDecryption Then
            CryptEx_Unprotect_Data = True
        Else
            CryptEx_Unprotect_Data = False
        End If
    End If
    
    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````'
    
'````````````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_VBProject:

    If Application.VBE.ActiveVBProject.Name = Encrypted_Data Then
         Set Obj_VBProject = Application.VBE.ActiveVBProject
         Current_ModuleName = Application.VBE.ActiveCodePane.CodeModule.Name
    Else
        For Each VBProject In Application.VBE.VBProjects
            If VBProject.Name = Encrypted_Data Then
                Set Obj_VBProject = VBProject: Exit For
            End If
        Next VBProject
    End If

    Count_VBComponent = Obj_VBProject.VBComponents.Count
    Crypt_ComponentName = Find_ModuleByGUID(GUID_VBComponent).Name
    
    For Each Obj_VBComponent In Obj_VBProject.VBComponents
        If Obj_VBComponent.Name <> Crypt_ComponentName Then
            If Len(Current_ModuleName) = 0 Then
                GoSub GS¦Decrypt_Component
            Else
                If Current_ModuleName <> Obj_VBComponent.Name Then
                    GoSub GS¦Decrypt_Component
                End If
            End If
        End If
    Next Obj_VBComponent
    
    Return
'````````````````````````````````````````````````````````````````````````````'
    
'`````````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_VBComponent:
    
    For Each VBComponent In Application.VBE.ActiveVBProject.VBComponents
        If VBComponent.Name = Encrypted_Data Then
            Set Obj_VBProject = Application.VBE.ActiveVBProject: Exit For
        End If
    Next

    If Obj_VBProject Is Nothing Then
        For Each VBProject In Application.VBE.VBProjects
            For Each VBComponent In VBProject.VBComponents
                If VBComponent.Name = Encrypted_Data Then
                    Set Obj_VBProject = VBProject: Exit For
                End If
            Next VBComponent
            If Not Obj_VBProject Is Nothing Then Exit For
        Next VBProject
    End If

    Set Obj_VBComponent = Obj_VBProject.VBComponents(Encrypted_Data)
    GoSub GS¦Decrypt_Component

    Return
'`````````````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦DeCryptEx_Procedure:

    If InStr(1, Encrypted_Data, ".") = 0 Then
        Error_Message = _
        "Проверьте входные данные на соответствие шаблону - ""Module.Procedure""" & _
        "Текущий шаблон (Encrypted_Data = " & Encrypted_Data & ")"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка определения процедуры")
    End If
    
    Module_Name = Split(Encrypted_Data, ".")(0): Procedure_Name = Split(Encrypted_Data, ".")(1)
    If Not IsModule_Exists(Module_Name) Then
        Error_Message = "Модуль """ & Module_Name & """ отсутствует в текущем проекте VBA!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска модуля")
        CryptEx_Unprotect_Data = False: Exit Function
    End If
    
    Set Obj_VBProject = ThisWorkbook.VBProject
    Set Obj_VBComponent = Obj_VBProject.VBComponents(Module_Name)
    Set Obj_CodeModule = Obj_VBComponent.CodeModule
    
    Set Encrypted_Lines = New Collection
    If Not Find_EncryptedBlock(Obj_CodeModule, Procedure_Name, GUID_VBComponent, _
                               Start_Line, End_Line, Encrypted_Lines) Then
        Error_Message = "Зашифрованная процедура (" & Procedure_Name & _
                                                 ") не найдена в модуле """ & Module_Name & """"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка поиска зашифрованной процедуры")
        CryptEx_Unprotect_Data = False: Exit Function
    End If
    
    GoSub GS¦Decrypt_Procedure
    
    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_File:

    Data_ToDecrypt = Conversion_FileToBytes(File_Path)

    If IsEmpty(Data_ToDecrypt) Then
        CryptEx_Unprotect_Data = Empty
        Return
    End If

    Data_Bytes = Data_ToDecrypt
    File_Details = Get_PathFileNameAndExtension(File_Path)

    If Not File_Details(2) = "vbfc" Then Return
 
    Inx_LB1 = LBound(Data_Bytes, 1)
    Inx_UB1 = UBound(Data_Bytes, 1)
    File_Details(2) = "."

    For N = Inx_UB1 - 7 To Inx_UB1
        If Data_Bytes(N) = 0 Then Exit For
        File_Details(2) = File_Details(2) & Chr$(Data_Bytes(N))
    Next N

    ReDim Preserve Data_Bytes(Inx_LB1 To Inx_UB1 - 8): Data_ToDecrypt = Data_Bytes
    GoSub GS¦Decrypt_Data: Data_Bytes = CryptEx_Unprotect_Data
    
    If UBound(Data_Bytes) = -1 Then
        Error_Message = "Файл: " & File_Path & " - не был дешифрован! " & _
                        "Проверьте ключ для дешифрации (" & Key_Crypt & ")"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка дешифрации файла")
        CryptEx_Unprotect_Data = False
        Return
    End If
    
    File_Details(0) = FS_SavePath
    File_Path = File_Details(0) & "\" & File_Details(1) & File_Details(2)
    
    If File_IsOpen(File_Path) Then
        Error_Message = "Копия файла в папке с обработанными файлами " & _
                        "на данным момент открыта и не может быть заменена!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу для замены")
        CryptEx_Unprotect_Data = False: Return
    End If
    
    CryptEx_Unprotect_Data = Conversion_BytesToFile(Data_Bytes, File_Path)

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_Component:

    Set Obj_CodeModule = Obj_VBComponent.CodeModule

    With Obj_CodeModule
        If .CountOfLines > 0 Then
            Data_ToDecrypt = .Lines(1, .CountOfLines)
            Data_ToDecrypt = Replace(Data_ToDecrypt, "' //", vbNullString)
            Data_ToDecrypt = Replace(Data_ToDecrypt, " ", vbNullString)
            GoSub GS¦Decrypt_Data: Data_ToDecrypt = CryptEx_Unprotect_Data
            If Not Data_ToDecrypt = vbNullString Then
                .DeleteLines 1, .CountOfLines
                .AddFromString Data_ToDecrypt
                CryptEx_Unprotect_Data = True
            Else
                Error_Message = "Компонент: " & Obj_VBComponent.Name & " - не был дешифрован! " & _
                        "Проверьте ключ для дешифрации (" & Key_Crypt & ")"
                Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка дешифрации файла")
                CryptEx_Unprotect_Data = False
            End If
        End If
    End With

    Set Obj_CodeModule = Nothing

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_Procedure:

    ReDim Temp_Array(1 To Encrypted_Lines.Count)
    
    For N = 1 To Encrypted_Lines.Count
        Temp_Array(N) = Encrypted_Lines(N)
    Next N
    
    Data_ToDecrypt = Join(Temp_Array, vbNullString)
    Data_ToDecrypt = Replace(Data_ToDecrypt, "' // ", vbNullString)
    
    GoSub GS¦Decrypt_Data
    
    If Not CryptEx_Unprotect_Data = vbNullString Then
        Obj_CodeModule.DeleteLines Start_Line, End_Line - Start_Line + 1
        Obj_CodeModule.InsertLines Start_Line, CryptEx_Unprotect_Data
        CryptEx_Unprotect_Data = True
    Else
        Error_Message = "Процедура: " & Procedure_Name & " в компоненте:" & Module_Name & _
                        " - не была дешифрована! " & _
                        "Проверьте ключ для дешифрации (" & Key_Crypt & ")"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка дешифрации файла")
        CryptEx_Unprotect_Data = False
    End If
    
    Return
'``````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦Decrypt_Data:

    CryptEx_Unprotect_Data = Crypt_Unprotect_Data( _
                             CStr(Data_ToDecrypt), Key_Crypt, Crypt_Method, PIM, Base_64)

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Check_SavePath:
    
    FS_SavePath = FS_SavePath & Environ$("USERPROFILE")
    FS_SavePath = FS_SavePath & FILE_SYSTEM_LOCAL_APP_DATA
    FS_SavePath = FS_SavePath & FILE_SYSTEM_SAVE_PATH
    FS_SavePath = FS_SavePath & FILE_SYSTEM_SAVE_FOLDER_DECRYPTED_FILES
    
    If Dir(FS_SavePath, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, FS_SavePath, ByVal 0&)
    End If

    Return
'```````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NonExistent_Directory:

    Error_Message = "Файл или папка не существует: " & Encrypted_Data
    Call Show_ErrorMessage_Immediate(Error_Message, _
                                    "Ошибка поиска нужного файла/директории", , True)
    CryptEx_Unprotect_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NullString:

    Error_Message = "Нет данных для обработки: Len String = 0"
    Call Show_ErrorMessage_Immediate(Error_Message, "Передана пустая строка")
    CryptEx_Unprotect_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'==================================================================================================================='
Public Function modRSA_PoolFunctions( _
                Optional ByVal Type_Function As RSA_PoolFunctions = GenerateRandom_AESKeys, _
                Optional ByVal KeyLength_ForGenerateKeys As modRSA_KeyLength = AES_256, _
                Optional ByRef AES_Key As String = vbNullString, _
                Optional ByRef RSA_Key As String = vbNullString _
       ) As Variant
'-------------------------------------------------------------------------------------------------------------------'
' // Documentation: In Process
'-------------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````'
    Dim PublicKey_Size     As Long, PrivateKey_Size   As Long
    Dim Public_Key()       As Byte, Private_Key()     As Byte
    Dim PublicKey_Bytes()  As Byte, Key_Length        As Long
    Dim Data_Bytes()       As Byte, Data_Len          As Long
    '````````````````````````````````````````````````````````'
    Dim Handle_Alg         As LongPtr, Key_Bytes()    As Byte
    Dim Handle_Key         As LongPtr, Char_Index     As Long
    '````````````````````````````````````````````````````````'
    Dim Private_KeyBytes() As Byte
    Dim Encrypted_Bytes()  As Byte, Decrypted_Bytes() As Byte
    Dim Encrypted_Len      As Long, Decrypted_Len     As Long
    '````````````````````````````````````````````````````````'
    Dim Error_Message As String, Status As Long, i    As Long
    '````````````````````````````````````````````````````````'
    Const Hex_Characters As String = "0123456789abcdef"
    '````````````````````````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````'
    Select Case Type_Function
        Case AESKey_ProtectWithPublicKey:    GoSub GS¦AESKey_Protect
        Case AESKey_UnprotectWithPrivateKey: GoSub GS¦AESKey_Unprotect
        Case GenerateRandom_AESKeys:         GoSub GS¦Generate_AESKeys
        Case GenerateRandom_RSAKeys:         GoSub GS¦Generate_RSAKeys
    End Select
    
    Exit Function
    '`````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Generate_AESKeys:

    Select Case KeyLength_ForGenerateKeys
        Case AES_256, AES_512
        Case Else
            Error_Message = "Выберите длину ключа из списка AES (AES_256 или AES_512)!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Выбрана неподходящая длина ключа!")
            Exit Function
    End Select
    
    KeyLength_ForGenerateKeys = KeyLength_ForGenerateKeys / 8
    ReDim Key_Bytes(0 To (KeyLength_ForGenerateKeys - 1))
    
    Status = BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr("RNG"), 0, 0)
    
    If Status = STATUS_SUCCESS Then
        Status = BCryptGenRandom(Handle_Alg, Key_Bytes(0), KeyLength_ForGenerateKeys, 0)
        BCryptCloseAlgorithmProvider Handle_Alg, 0
    Else
        ' // < ToDo: Заменить на более безопасную VBA реализацию >
        ReDim Key_Bytes(0 To KeyLength_ForGenerateKeys - 1) As Byte: Randomize

        For i = 0 To KeyLength_ForGenerateKeys - 1
            Char_Index = Int((Len(Hex_Characters) * Rnd) + 1)
            Key_Bytes(i) = Asc(Mid$(Hex_Characters, Char_Index, 1))
        Next i
    End If
    
    modRSA_PoolFunctions = Base64_Encode_CryptoNG(Key_Bytes)
    
    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Generate_RSAKeys:

    Select Case KeyLength_ForGenerateKeys
        Case RSA_2048, RSA_3072, RSA_4096, RSA_8192
        Case Else
            Error_Message = "Выберите длину ключа из списка RSA (RSA_2048 - RSA_8192)!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Выбрана неподходящая длина ключа!")
            Exit Function
    End Select

    Key_Length = KeyLength_ForGenerateKeys
    ReDim Array_Keys(0 To 1) As String

    Status = BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr("RSA"), 0, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка открытия алгоритма RSA"): Exit Function
        
    End If

    Status = BCryptGenerateKeyPair(Handle_Alg, Handle_Key, Key_Length, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка генерации пары ключей RSA")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    Status = BCryptFinalizeKeyPair(Handle_Key, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка финализации ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    Status = BCryptExportKey(Handle_Key, 0, StrPtr("RSAPUBLICBLOB"), 0, 0, PublicKey_Size, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка получения размера публичного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    ReDim Public_Key(PublicKey_Size - 1)
    Status = BCryptExportKeyToBuffer(Handle_Key, 0, StrPtr("RSAPUBLICBLOB"), Public_Key(0), _
                                     PublicKey_Size, PublicKey_Size, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка экспорта публичного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    Status = BCryptExportKey(Handle_Key, 0, StrPtr("RSAFULLPRIVATEBLOB"), 0, 0, PrivateKey_Size, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка получения размера приватного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    ReDim Private_Key(PrivateKey_Size - 1)
    Status = BCryptExportKeyToBuffer(Handle_Key, 0, StrPtr("RSAFULLPRIVATEBLOB"), _
                                     Private_Key(0), PrivateKey_Size, PrivateKey_Size, 0)
                                     
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка экспорта приватного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    Array_Keys(0) = Base64_Encode_CryptoNG(Public_Key)
    Array_Keys(1) = Base64_Encode_CryptoNG(Private_Key)
    
    GoSub GS¦Clearing_Memory: modRSA_PoolFunctions = Array_Keys

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦AESKey_Protect:

    Status = BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr("RSA"), 0, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка открытия алгоритма"): Exit Function
    End If
    
    PublicKey_Bytes = Base64_Decode_CryptoNG(RSA_Key)
    Status = BCryptImportKeyPair(Handle_Alg, 0, StrPtr("RSAPUBLICBLOB"), Handle_Key, _
                                 PublicKey_Bytes(0), UBound(PublicKey_Bytes) + 1, 0)
                                 
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка импорта публичного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    Data_Bytes = StrConv(AES_Key, vbFromUnicode)
    Data_Len = UBound(Data_Bytes) + 1
    
    Status = BCryptEncrypt(Handle_Key, Data_Bytes(0), Data_Len, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, _
                           Encrypted_Len, BCRYPT_PAD_PKCS1)
                           
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка получения размера шифрования")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    ReDim Encrypted_Bytes(Encrypted_Len - 1)
    Status = BCryptEncryptToBuffer(Handle_Key, Data_Bytes(0), Data_Len, ByVal 0&, ByVal 0&, ByVal 0&, _
                                   Encrypted_Bytes(0), Encrypted_Len, Encrypted_Len, BCRYPT_PAD_PKCS1)
                                   
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка шифрования данных")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    GoSub GS¦Clearing_Memory: modRSA_PoolFunctions = Base64_Encode_CryptoNG(Encrypted_Bytes)
    
    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦AESKey_Unprotect:

    Status = BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr("RSA"), 0, 0)
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка открытия алгоритма"): Exit Function
    End If

    Private_KeyBytes = Base64_Decode_CryptoNG(RSA_Key)
    Status = BCryptImportKeyPair(Handle_Alg, 0, StrPtr("RSAFULLPRIVATEBLOB"), Handle_Key, _
                                 Private_KeyBytes(0), UBound(Private_KeyBytes) + 1, 0)
                                 
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка импорта приватного ключа")
        GoSub GS¦Clearing_Memory: Exit Function
    End If

    Encrypted_Bytes = Base64_Decode_CryptoNG(AES_Key)
    Encrypted_Len = UBound(Encrypted_Bytes) + 1
    Status = BCryptDecrypt(Handle_Key, Encrypted_Bytes(0), Encrypted_Len, ByVal 0&, ByVal 0&, ByVal 0&, _
                           ByVal 0&, ByVal 0&, Decrypted_Len, BCRYPT_PAD_PKCS1)
                           
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка получения размера дешифрования")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    ReDim Decrypted_Bytes(Decrypted_Len - 1)
    Status = BCryptDecryptToBuffer(Handle_Key, Encrypted_Bytes(0), Encrypted_Len, ByVal 0&, ByVal 0&, ByVal 0&, _
                                   Decrypted_Bytes(0), Decrypted_Len, Decrypted_Len, BCRYPT_PAD_PKCS1)
                                   
    If Status <> STATUS_SUCCESS Then
        Error_Message = "Обнаружена проблема на уровне вызова WinAPI! Сообщите разработчику о выявленной проблеме!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Ошибка дешифрования данных")
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    
    ReDim Preserve Decrypted_Bytes(Decrypted_Len - 1)
    GoSub GS¦Clearing_Memory: modRSA_PoolFunctions = StrConv(Decrypted_Bytes, vbUnicode)

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    
'````````````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Key) Then BCryptDestroyKey Handle_Key
    If CBool(Handle_Alg) Then BCryptCloseAlgorithmProvider Handle_Alg, 0&

    Return
'````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'===================================================================================='
Public Sub Open_FolderWithProcessedFiles()
'------------------------------------------------------------------------------------'
' // Documentation: In Process
'------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````'
    Dim Folder_ProcessedFiles As String, Obj_Shell As Object
    '```````````````````````````````````````````````````````'
    Const SW_SHOWNORMAL As Long = 1&
    '```````````````````````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````````````'
    Folder_ProcessedFiles = Folder_ProcessedFiles & Environ$("USERPROFILE")
    Folder_ProcessedFiles = Folder_ProcessedFiles & FILE_SYSTEM_LOCAL_APP_DATA
    Folder_ProcessedFiles = Folder_ProcessedFiles & FILE_SYSTEM_SAVE_PATH
    '`````````````````````````````````````````````````````````````````````````'
    
    '``````````````````````````````````````````````````````````````````````'
    If Dir(Folder_ProcessedFiles, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, Folder_ProcessedFiles, ByVal 0&)
    End If
    '``````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    Call ShellExecute(ByVal 0&, "Open", Folder_ProcessedFiles, "", "", SW_SHOWNORMAL)
    '````````````````````````````````````````````````````````````````````````````````'
    
'------------------------------------------------------------------------------------'
End Sub
'===================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'======================================================================'
Private Sub Init_VBD_Kit_Cryptography()
'----------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_Undefined Then
        Call Get_MSOffice_Type_Building
    End If
    '``````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Application = Type_MSOffice_App_Undefined Then
        Call Get_MSOffice_Type_Application
    End If
    '``````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------'
End Sub
'======================================================================'


'=============================================================================================='
Private Sub Init_AccessObjectModel( _
            Optional ByVal Type_Document As MSOffice_Type_Document_Group = mso_Excel, _
            Optional ByVal Macro_ForcedChange As Boolean = True _
        )
'----------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````'
    Select Case Glb_MSOffice_Type_Application
        Case Type_MSOffice_Access, Type_MSOffice_Outlook: Exit Sub
    End Select
    '`````````````````````````````````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    On Error Resume Next
    
    If Not Len(Application.VBE.ActiveVBProject.Name) = 0 Then
        If Err.Number = 0 Then On Error GoTo 0: Exit Sub
    End If
    
    On Error GoTo 0
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    Call MacroSecurity_MacroParameters_Assign( _
                Type_Document, Without_PrivilegesChanges, Access_Provided, Macro_ForcedChange _
         )
    '``````````````````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------'
End Sub
'=============================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================'
Private Function Get_MSOffice_Type_Building() As MSOffice_Type_Building
'------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````'
    Dim Exl_App As Object, Flag_MSOffice_Type As Boolean
    '```````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    #If Has_PtrSafe Then
        #If Win64 Then
            GoSub GS¦Check_MSOffice_Support: GoSub GS¦Check_MSOffice_2019_365
            If Flag_MSOffice_Type Then
                Glb_MSOffice_Type_Building = Type_MSOffice_2019_365_64Bit
                Get_MSOffice_Type_Building = Type_MSOffice_2019_365_64Bit
            Else
                Glb_MSOffice_Type_Building = Type_MSOffice_2016_64Bit
                Get_MSOffice_Type_Building = Type_MSOffice_2016_64Bit
            End If
        #Else
            GoSub GS¦Check_MSOffice_Support: GoSub GS¦Check_MSOffice_2019_365
            If Flag_MSOffice_Type Then
                Glb_MSOffice_Type_Building = Type_MSOffice_2019_365_32Bit
                Get_MSOffice_Type_Building = Type_MSOffice_2019_365_32Bit
            Else
                Glb_MSOffice_Type_Building = Type_MSOffice_2016_32Bit
                Get_MSOffice_Type_Building = Type_MSOffice_2016_32Bit
            End If
        #End If
    #Else
        GoSub GS¦Support_ErrorMessage_MSOffice2010
        Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport
        Get_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport
    #End If

    Exit Function
    '````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````'
GS¦Check_MSOffice_Support:

    If CLng(Val(Application.Version)) < 16 Then

        GoSub GS¦Support_ErrorMessage_MSOffice2010
        Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport
        Get_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport

        Exit Function

    End If

    Return
'```````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````'
GS¦Check_MSOffice_2019_365:

    Flag_MSOffice_Type = False

    If Application.Name = "Microsoft Excel" Then
        Set Exl_App = Application
    Else
        Set Exl_App = CreateObject("Excel.Application")
    End If

    If CLng(Left(Exl_App.CalculationVersion, 2)) = 19 Then Flag_MSOffice_Type = True

    Set Exl_App = Nothing

    Return
'````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Support_ErrorMessage_MSOffice2010:

    MsgBox "Данный программный модуль ""не тащит"" за собой " & _
           "поддержку более старого пакета MS Office (до 2016 года)!" _
           & vbNewLine & "Обновите свой MS Office, чтобы пользоваться функционалом данного компонента!", _
                                                                      vbInformation, "[Cyber_Automation]"

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'============================================================================================================'
Private Function Get_MSOffice_Type_Application() As MSOffice_Type_Application
'------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Select Case Application.Name

        Case "Microsoft Excel"
            Glb_MSOffice_Type_Application = Type_MSOffice_Excel
            Get_MSOffice_Type_Application = Type_MSOffice_Excel

        Case "Microsoft Word"
            Glb_MSOffice_Type_Application = Type_MSOffice_Word
            Get_MSOffice_Type_Application = Type_MSOffice_Word

        Case "Microsoft PowerPoint"
            Glb_MSOffice_Type_Application = Type_MSOffice_PowerPoint
            Get_MSOffice_Type_Application = Type_MSOffice_PowerPoint

        Case "Microsoft Access"
            Glb_MSOffice_Type_Application = Type_MSOffice_Access
            Get_MSOffice_Type_Application = Type_MSOffice_Access

        Case "Microsoft Outlook"
            Glb_MSOffice_Type_Application = Type_MSOffice_Outlook
            Get_MSOffice_Type_Application = Type_MSOffice_Outlook

        Case Else
            GoSub GS¦Support_ErrorMessage_Application
            Glb_MSOffice_Type_Application = Type_MSOffice_App_NotSupport
            Get_MSOffice_Type_Application = Type_MSOffice_App_NotSupport

    End Select

    Exit Function
    '````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Support_ErrorMessage_Application:

    MsgBox "Данный программный модуль не поддерживает текущий программный продукт: " & Application.Name _
           & vbNewLine & "Обратитесь к автору данного компонента для возможного внедрения совместимости!", _
                                                                      vbInformation, "[Cyber_Automation]"

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'================================================================================================================='
Private Function Get_ID_OperationSystem( _
                 Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen _
        ) As String
'-----------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````'
    Dim System_Type  As String
    Dim Device_Name  As String
    Dim Product_Key  As String
    Dim Drive_Letter As String
    '`````````````````````````````````'
    Dim Intermediate_ID()     As Byte
    Dim Drive_SerialNumber  As String
    Dim Install_Date        As String
    '`````````````````````````````````'
    Static Obj_WMIService   As Object
    Static Obj_FSO          As Object
    Static Obj_Drive        As Object
    '`````````````````````````````````'
    Dim Obj_OperatingSystem As Variant
    '`````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    If Obj_WMIService Is Nothing Then Set Obj_WMIService = GetObject("winmgmts:\\.\root\CIMV2")
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    For Each Obj_OperatingSystem In Obj_WMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

        System_Type = Environ("OS") & ": " & Obj_OperatingSystem.OSArchitecture
        Device_Name = Obj_OperatingSystem.CSName

        If Not IsNull(Obj_OperatingSystem.SerialNumber) Then
            Product_Key = Obj_OperatingSystem.SerialNumber
        Else
            Product_Key = "00000-00000-00000-00000"
        End If

        GoSub GS¦Get_Drive_SerialNumber: Install_Date = Obj_OperatingSystem.InstallDate

        Intermediate_ID = System_Type & "|" & Device_Name & "|" & _
                          Product_Key & "|" & Drive_SerialNumber & "|" & Install_Date

    Next Obj_OperatingSystem
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    Select Case Crypt_Method
        Case Crypt_VBACode_XOR, Crypt_VBACode_AES_256
            Get_ID_OperationSystem = Hashing_NativeVBACode(Intermediate_ID)

        Case Crypt_WinAPI_AES_CryptoAPI
            Get_ID_OperationSystem = Hashing_CryptoAPI(CStr(Intermediate_ID))

        Case Crypt_WinAPI_AES_CryptoNextGen
            Get_ID_OperationSystem = Hashing_CryptoNG(Intermediate_ID)
    End Select

    Exit Function
    '````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Drive_SerialNumber:

    If Obj_FSO Is Nothing Then Set Obj_FSO = CreateObject("Scripting.FileSystemObject")

    If Obj_Drive Is Nothing Then
        Drive_Letter = Left$(Environ$("SystemRoot"), 1)
        Set Obj_Drive = Obj_FSO.GetDrive(Obj_FSO.GetDriveName(Drive_Letter & ":\"))
    End If

    If Not Obj_Drive Is Nothing Then
        Drive_SerialNumber = Obj_Drive.SerialNumber
    Else
        Drive_SerialNumber = "Unknown"
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================='


'================================================================================================================='
Private Function Get_ID_Hardware( _
                 Optional ByVal Crypt_Method As VBProject_Crypt_Algorithm = Crypt_WinAPI_AES_CryptoNextGen, _
                 Optional ByVal Type_Hardware As VBProject_Type_Hardware = Type_Win32_BaseBoard _
        ) As String
'-----------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````'
    Static Obj_WMIService As Object
    Dim Obj_Item          As Variant
    Dim Intermediate_ID   As String
    Dim Byte_Data()       As Byte
    '```````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    If Obj_WMIService Is Nothing Then Set Obj_WMIService = GetObject("winmgmts:\\.\root\CIMV2")
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Select Case Type_Hardware
        Case Type_Win32_BaseBoard:       GoSub GS¦Get_Hardware_BaseBoard
        Case Type_Win32_Processor:       GoSub GS¦Get_Hardware_Processor
        Case Type_Win32_VideoController: GoSub GS¦Get_Hardware_VideoController
        Case Type_Win32_PhysicalMemory:  GoSub GS¦Get_Hardware_PhysicalMemory
        Case Type_Win32_Firmware:        GoSub GS¦Get_Hardware_Firmware
    End Select
    '`````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````'
    Select Case Crypt_Method
        Case Crypt_VBACode_XOR, Crypt_VBACode_AES_256
            Byte_Data = Intermediate_ID
            Get_ID_Hardware = Hashing_NativeVBACode(Byte_Data)

        Case Crypt_WinAPI_AES_CryptoAPI
            Get_ID_Hardware = Hashing_CryptoAPI(Intermediate_ID)

        Case Crypt_WinAPI_AES_CryptoNextGen
            Byte_Data = Intermediate_ID
            Get_ID_Hardware = Hashing_CryptoNG(Byte_Data)
    End Select

    Exit Function
    '````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Hardware_BaseBoard:

    For Each Obj_Item In Obj_WMIService.ExecQuery("SELECT * FROM Win32_BaseBoard")
        With Obj_Item
            Intermediate_ID = .Manufacturer & "|" & _
                              .Product & "|" & _
                              .SerialNumber
        End With
    Next Obj_Item

    Set Obj_Item = Nothing

    Return
'`````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Hardware_Processor:

    For Each Obj_Item In Obj_WMIService.ExecQuery("SELECT * FROM Win32_Processor")
        With Obj_Item
            Intermediate_ID = Intermediate_ID & .Name & "|" & _
                             .Manufacturer & "|" & _
                             .Description & "|" & _
                             .MaxClockSpeed & "|" & _
                             .NumberOfCores & "|" & _
                             .NumberOfLogicalProcessors & "|"
        End With
    Next Obj_Item

    Set Obj_Item = Nothing

    Return
'`````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Hardware_VideoController:

    For Each Obj_Item In Obj_WMIService.ExecQuery("SELECT * FROM Win32_VideoController")
        With Obj_Item
            Intermediate_ID = Intermediate_ID & .Caption & "|" & _
                              .AdapterCompatibility & "|" & _
                              .VideoProcessor & "|"
        End With
    Next Obj_Item

    Set Obj_Item = Nothing

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Hardware_PhysicalMemory:

    For Each Obj_Item In Obj_WMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
        With Obj_Item
            Intermediate_ID = Intermediate_ID & .Manufacturer & "|" & _
                              Round(.Capacity / 1024& _
                                               / 1024& _
                                               / 1024&, 2&) & "|" & _
                             .Speed & "|"
        End With
    Next Obj_Item

    Set Obj_Item = Nothing

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````'
GS¦Get_Hardware_Firmware:

    For Each Obj_Item In Obj_WMIService.ExecQuery("SELECT * FROM Win32_BIOS")
        With Obj_Item
            Intermediate_ID = .Manufacturer & "|" & .SerialNumber
        End With
    Next Obj_Item

    Set Obj_Item = Nothing

    Return
'`````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'================================================================================================================'
Private Function Hashing_CryptoNG( _
                 ByRef Data_Bytes() As Byte, _
                 Optional ByRef Hashing_Algorithm As String = "SHA1" _
        ) As String
'----------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Return_HashSumm As String, Error_Message As String, Position As Long, Inx_LB1   As Long
    Dim Ptr_Data As LongPtr, Handle_Hash As LongPtr, Handle_Alg As LongPtr, Size_Data   As Long, Inx_UB1 As Long
    Dim Buffer_Hash() As Byte, Buffer_HashObject() As Byte, Length_Hash As Long, Length As Long, i       As Long
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````'
    Inx_LB1 = LBound(Data_Bytes, 1)
    Inx_UB1 = UBound(Data_Bytes, 1)

    Size_Data = Inx_UB1 - Inx_LB1 + 1
    '````````````````````````````````'

    '````````````````````````````````````'
    Ptr_Data = VarPtr(Data_Bytes(Inx_LB1))
    '````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    If BCryptOpenAlgorithmProvider(Handle_Alg, _
                                   StrPtr(Hashing_Algorithm & vbNullChar), 0, 0) Then
        Error_Message = "BCryptOpenAlgorithmProvider: Ошибка доступа к провайдеру: " _
                                                                    & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function

    End If
    '``````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    If BCryptGetProperty(Handle_Alg, _
                         StrPtr("ObjectLength" & vbNullString), Length, LenB(Length), 0, 0) <> 0 Then
        Error_Message = "BCryptGetProperty: Ошибка получения длины объекта хэша: " _
                                                                  & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    ReDim Buffer_HashObject(0 To Length - 1)

    If BCryptGetProperty(Handle_Alg, _
                         StrPtr("HashDigestLength" & vbNullChar), Length_Hash, LenB(Length_Hash), 0, 0) <> 0 Then
        Error_Message = "BCryptGetProperty: Ошибка получения итоговой длины хэша: " _
                                                                   & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    ReDim Buffer_Hash(0 To Length_Hash - 1)

    If BCryptCreateHash(Handle_Alg, Handle_Hash, Buffer_HashObject(0), Length, 0, 0, 0) <> 0 Then
        Error_Message = "BCryptCreateHash: Ошибка создания объекта хэша: " _
                                                          & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    If BCryptHashData(Handle_Hash, ByVal Ptr_Data, Size_Data) <> 0 Then
        Error_Message = "BCryptHashData: Ошибка добавления данных для хеширования: " _
                                                                    & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    If BCryptFinishHash(Handle_Hash, Buffer_Hash(0), Length_Hash, 0) <> 0 Then
        Error_Message = "BCryptFinishHash: Ошибка извлечения значения хэша: " _
                                                           & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoNG = vbNullString: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Inx_LB1 = LBound(Buffer_Hash, 1)
    Inx_UB1 = UBound(Buffer_Hash, 1)

    GoSub GS¦Clearing_Memory
    '```````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    Return_HashSumm = String$((Inx_UB1 - Inx_LB1 + 1) * 2, vbNullChar): Position = 1

    For i = Inx_LB1 To Inx_UB1
        Mid$(Return_HashSumm, Position) = Right$("00" & Hex(Buffer_Hash(i)), 2)
        Position = Position + 2
    Next i

    Hashing_CryptoNG = LCase$(Return_HashSumm)

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````'

''``````````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Hash) Then BCryptDestroyHash Handle_Hash
    If CBool(Handle_Alg) Then BCryptCloseAlgorithmProvider Handle_Alg, 0

    Return
''``````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================'


'=================================================================================================='
Private Function Hashing_CryptoAPI(ByVal Source_Text As String) As String
'--------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim Handle_Prov  As LongPtr
    Dim Handle_Hash  As LongPtr
    Dim Hash_Len     As Long
    Dim Hash_Value() As Byte, Inx_LB1 As Long
    Dim Data_Bytes() As Byte, Inx_UB1 As Long
    Dim i As Long, Return_HashSumm  As String
    Dim Error_Message As String
    '``````````````````````````````````````````'

    '``````````````````````````````````````````'
    Data_Bytes = StrConv(Source_Text, vbUnicode)
    '``````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If CryptAcquireContext(Handle_Prov, _
                           vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        Error_Message = "CryptAcquireContext: Ошибка при получении контекста криптопровайдера: " _
                                                                                & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoAPI = vbNullString: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    If CryptCreateHash(Handle_Prov, CALG_SHA1, 0, 0, Handle_Hash) = 0 Then
        Error_Message = "CryptCreateHash: Ошибка при создании объекта хэша: " _
                                                             & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoAPI = vbNullString: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    If CryptHashData(Handle_Hash, _
                     Data_Bytes(0), UBound(Data_Bytes) + 1, 0) = 0 Then
        Error_Message = "CryptHashData: Ошибка при вычислении хэша: " _
                                                     & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoAPI = vbNullString: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    If CryptGetHashParam(Handle_Hash, HP_HASHSIZE, Hash_Len, 4, 0) = 0 Then
        Error_Message = "CryptGetHashParam: Ошибка при получении размера хэша: " _
                                                                & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoAPI = vbNullString: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````'
    ReDim Hash_Value(Hash_Len - 1)
    '`````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    If CryptGetHashParam(Handle_Hash, _
                         HP_HASHVAL, Hash_Value(0), Hash_Len, 0) = 0 Then
        Error_Message = "CryptGetHashParam: Ошибка при получении значения хэша: " _
                                                                 & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_CryptoAPI = vbNullString: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````'
    Inx_LB1 = LBound(Hash_Value, 1)
    Inx_UB1 = UBound(Hash_Value, 1)

    GoSub GS¦Clearing_Memory
    '``````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    For i = Inx_LB1 To Inx_UB1
        Return_HashSumm = Return_HashSumm & Right$("0" & Hex$(Hash_Value(i)), 2)
    Next i

    Hashing_CryptoAPI = LCase$(Return_HashSumm)

    Exit Function
    '``````````````````````````````````````````````````````````````````````````'

''````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Hash) Then CryptDestroyHash Handle_Hash
    If CBool(Handle_Prov) Then CryptReleaseContext Handle_Prov, 0

    Return
''````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'======================================================='
Private Function Hashing_NativeVBACode( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'-------------------------------------------------------'
    Hashing_NativeVBACode = SHA1_HexDefault(Data_Bytes)
'-------------------------------------------------------'
End Function
'======================================================='


'=================================================================='
Private Function SHA1_HexDefault( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'------------------------------------------------------------------'
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long

    SHA1_Default Data_Bytes, H1, H2, H3, H4, H5
    SHA1_HexDefault = SHA1_DecToHex5(H1, H2, H3, H4, H5)
'------------------------------------------------------------------'
End Function
'=================================================================='


'=================================================================='
Private Function SHA1_Hex( _
                 ByRef message() As Byte, _
                 ByRef Key1 As Long, _
                 ByRef Key2 As Long, _
                 ByRef Key3 As Long, _
                 ByRef Key4 As Long _
        ) As String
'------------------------------------------------------------------'
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long

    SHA1_x message, Key1, Key2, Key3, Key4, H1, H2, H3, H4, H5
    SHA1_Hex = SHA1_DecToHex5(H1, H2, H3, H4, H5)
'------------------------------------------------------------------'
End Function
'=================================================================='


'======================================================================================'
Private Sub SHA1_Default( _
            ByRef message() As Byte, _
            ByRef H1 As Long, _
            ByRef H2 As Long, _
            ByRef H3 As Long, _
            ByRef H4 As Long, _
            ByRef H5 As Long _
        )
'--------------------------------------------------------------------------------------'
    SHA1_x message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, H1, H2, H3, H4, H5
'--------------------------------------------------------------------------------------'
End Sub
'======================================================================================'


'========================================================================================================================================'
Private Sub SHA1_x( _
            ByRef message() As Byte, _
            ByVal Key1 As Long, _
            ByVal Key2 As Long, _
            ByVal Key3 As Long, _
            ByVal Key4 As Long, _
            ByRef H1 As Long, _
            ByRef H2 As Long, _
            ByRef H3 As Long, _
            ByRef H4 As Long, _
            ByRef H5 As Long _
        )
'----------------------------------------------------------------------------------------------------------------------------------------'
     Dim U  As Long, P As Long
     Dim FB As SHA1_FourBytes, OL As SHA1_OneLong
     Dim i  As Integer
     Dim W(80) As Long
     Dim A As Long, B As Long, C As Long
     Dim D As Long, E As Long, T As Long

     H1 = &H67452301: H2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0
     U = UBound(message) + 1: OL.L = SHA1_U32ShiftLeft3(U): A = U \ &H20000000: LSet FB = OL

     ReDim Preserve message(0 To (U + 8 And -64) + 63): message(U) = 128

     U = UBound(message)
     message(U - 4) = A
     message(U - 3) = FB.D
     message(U - 2) = FB.C
     message(U - 1) = FB.B
     message(U) = FB.A

     While P < U
         For i = 0 To 15
             FB.D = message(P)
             FB.C = message(P + 1)
             FB.B = message(P + 2)
             FB.A = message(P + 3)
             LSet OL = FB
             W(i) = OL.L
             P = P + 4
         Next i

         For i = 16 To 79
             W(i) = SHA1_U32RotateLeft1(W(i - 3) Xor W(i - 8) Xor W(i - 14) Xor W(i - 16))
         Next i

         A = H1: B = H2: C = H3: D = H4: E = H5

         For i = 0 To 19
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key1), ((B And C) Or ((Not B) And D)))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i
         For i = 20 To 39
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key2), (B Xor C Xor D))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i
         For i = 40 To 59
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key3), ((B And C) Or _
                                                                                                            (B And D) Or (C And D)))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i
         For i = 60 To 79
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key4), (B Xor C Xor D))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i

         H1 = SHA1_U32Add(H1, A): H2 = SHA1_U32Add(H2, B): H3 = SHA1_U32Add(H3, C): H4 = SHA1_U32Add(H4, D): H5 = SHA1_U32Add(H5, E)
     Wend
'----------------------------------------------------------------------------------------------------------------------------------------'
End Sub
'========================================================================================================================================'


'=========================================================='
Private Function SHA1_U32Add( _
                 ByVal A As Long, _
                 ByVal B As Long _
        ) As Long
'----------------------------------------------------------'
    If (A Xor B) < 0 Then
        SHA1_U32Add = A + B
    Else
        SHA1_U32Add = (A Xor &H80000000) + B Xor &H80000000
    End If
'----------------------------------------------------------'
End Function
'=========================================================='


'================================================================================='
Private Function SHA1_U32ShiftLeft3( _
                 ByVal A As Long _
        ) As Long
'---------------------------------------------------------------------------------'
    SHA1_U32ShiftLeft3 = (A And &HFFFFFFF) * 8
    If A And &H10000000 Then SHA1_U32ShiftLeft3 = SHA1_U32ShiftLeft3 Or &H80000000
'---------------------------------------------------------------------------------'
End Function
'================================================================================='


'==================================================================================='
Private Function SHA1_U32RotateLeft1( _
                 ByVal A As Long _
        ) As Long
'-----------------------------------------------------------------------------------'
    SHA1_U32RotateLeft1 = (A And &H3FFFFFFF) * 2
    If A And &H40000000 Then SHA1_U32RotateLeft1 = SHA1_U32RotateLeft1 Or &H80000000
    If A And &H80000000 Then SHA1_U32RotateLeft1 = SHA1_U32RotateLeft1 Or 1
'-----------------------------------------------------------------------------------'
End Function
'==================================================================================='


'========================================================================================'
Private Function SHA1_U32RotateLeft5( _
                 ByVal A As Long _
        ) As Long
'----------------------------------------------------------------------------------------'
    SHA1_U32RotateLeft5 = (A And &H3FFFFFF) * 32 Or (A And &HF8000000) \ &H8000000 And 31
    If A And &H4000000 Then SHA1_U32RotateLeft5 = SHA1_U32RotateLeft5 Or &H80000000
'----------------------------------------------------------------------------------------'
End Function
'========================================================================================'


'======================================================================================'
Private Function SHA1_U32RotateLeft30( _
                 ByVal A As Long _
        ) As Long
'--------------------------------------------------------------------------------------'
    SHA1_U32RotateLeft30 = (A And 1) * &H40000000 Or (A And &HFFFC) \ 4 And &H3FFFFFFF
    If A And 2 Then SHA1_U32RotateLeft30 = SHA1_U32RotateLeft30 Or &H80000000
'--------------------------------------------------------------------------------------'
End Function
'======================================================================================'


'==============================================================='
Private Function SHA1_DecToHex5( _
                 ByVal H1 As Long, _
                 ByVal H2 As Long, _
                 ByVal H3 As Long, _
                 ByVal H4 As Long, _
                 ByVal H5 As Long _
        ) As String
'---------------------------------------------------------------'
    Dim H As String, L As Long
    SHA1_DecToHex5 = "0000000000000000000000000000000000000000"
    H = Hex(H1): L = Len(H): Mid(SHA1_DecToHex5, 9 - L, L) = H
    H = Hex(H2): L = Len(H): Mid(SHA1_DecToHex5, 17 - L, L) = H
    H = Hex(H3): L = Len(H): Mid(SHA1_DecToHex5, 25 - L, L) = H
    H = Hex(H4): L = Len(H): Mid(SHA1_DecToHex5, 33 - L, L) = H
    H = Hex(H5): L = Len(H): Mid(SHA1_DecToHex5, 41 - L, L) = H
    SHA1_DecToHex5 = LCase$(SHA1_DecToHex5)
'---------------------------------------------------------------'
End Function
'==============================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'================================================================='
Private Function AES_128_Encrypt_CryptoNG( _
                 ByRef Input_String As String, _
                 ByRef Input_Key As String _
        ) As String
'-----------------------------------------------------------------'

    '``````````````````````````````````````'
    Dim Data_Bytes() As Byte, Key() As Byte
    '``````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    Data_Bytes = Input_String: Key = Input_Key
    AES_128_Encrypt_CryptoNG = NextGen_EncryptData(Data_Bytes, Key)
    '`````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------'
End Function
'================================================================='


'========================================================='
Private Function AES_128_Decrypt_CryptoNG( _
                 ByRef Encrypted_String As String, _
                 ByRef Input_Key As String _
        ) As String
'---------------------------------------------------------'

    '`````````````````````````````````````````````````````'
    Dim Data_Bytes() As Byte, Key() As Byte, Out() As Byte
    '`````````````````````````````````````````````````````'

    '`````````````````````````````````````````````'
    Data_Bytes = Encrypted_String: Key = Input_Key
    Call NextGen_DecryptData(Data_Bytes, Key, Out)

    AES_128_Decrypt_CryptoNG = Out
    '`````````````````````````````````````````````'

'---------------------------------------------------------'
End Function
'========================================================='


'======================================================================================='
Private Function NextGen_EncryptData( _
                 ByRef Input_Data() As Byte, _
                 ByRef Input_Key() As Byte _
        ) As Byte()
'---------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Key_Hash()  As Byte
    Dim Data_Hash() As Byte
    Dim Data_Length As Long
    '```````````````````````````'
    Dim Encrypt_Bytes()  As Byte
    Dim IV(0 To 15)      As Byte
    Dim Encrypted_Data() As Byte
    '```````````````````````````'

    '````````````````````````````````````````````````````````````'
    Key_Hash = NextGen_HashGeneration_Wrapper(Input_Key, "SHA1")
    Data_Hash = NextGen_HashGeneration_Wrapper(Input_Data, "SHA1")
    '````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Data_Length = UBound(Input_Data) - LBound(Input_Data) + 1
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    ReDim Encrypt_Bytes(0 To Data_Length + 23)
    RtlMoveMemory Encrypt_Bytes(0), Data_Length, 4
    RtlMoveMemory Encrypt_Bytes(4), Input_Data(LBound(Input_Data)), Data_Length
    RtlMoveMemory Encrypt_Bytes(Data_Length + 4), Data_Hash(0), 20
    '``````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    NextGen_Random_Wrapper IV
    Encrypted_Data = NextGen_Encrypt(VarPtr(Encrypt_Bytes(0)), _
                             Data_Length + 24, VarPtr(IV(0)), VarPtr(Key_Hash(0)), 16)
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Erase Encrypt_Bytes
    ReDim Preserve Encrypted_Data(LBound(Encrypted_Data) To UBound(Encrypted_Data) + 16)
    RtlMoveMemory Encrypted_Data(UBound(Encrypted_Data) - 15), IV(0), 16
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````'
    NextGen_EncryptData = Encrypted_Data
    '```````````````````````````````````'

'---------------------------------------------------------------------------------------'
End Function
'======================================================================================='


'======================================================================================'
Private Function NextGen_DecryptData( _
                 ByRef Input_Data() As Byte, _
                 ByRef Input_Key() As Byte, _
                 ByRef Out_Decrypted() As Byte _
        ) As Boolean
'--------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Array_Length As Long
    Dim Key_Hash()   As Byte
    Dim Ptr_IV       As LongPtr
    Dim Decrypted_Data() As Byte
    Dim Data_Length      As Long
    Dim Hash_Result()    As Byte
    Dim L As Byte
    '```````````````````````````'

    '````````````````````````````````````````````'
    If LBound(Input_Data) <> 0 Then Exit Function
    '````````````````````````````````````````````'

    '``````````````````````````````````````'
    Array_Length = UBound(Input_Data) + 1
    If Array_Length < 20 Then Exit Function
    '``````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Key_Hash = NextGen_HashGeneration_Wrapper(Input_Key, "SHA1")
    '``````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    Ptr_IV = VarPtr(Input_Data(UBound(Input_Data) - 15))
    Decrypted_Data = NextGen_Decrypt(VarPtr(Input_Data(0)), _
                             UBound(Input_Data) - LBound(Input_Data) - 15, _
                             Ptr_IV, VarPtr(Key_Hash(0)), 16)
    '````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    If StrPtr(Decrypted_Data) = 0 Then Exit Function
    If UBound(Decrypted_Data) < 3 Then Exit Function
    '```````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    RtlMoveMemory Data_Length, Decrypted_Data(0), 4
    If Data_Length > (UBound(Decrypted_Data) - 3) Or Data_Length < 0 Then Exit Function
    Hash_Result = NextGen_HashGeneration(VarPtr(Decrypted_Data(4)), Data_Length, "SHA1")
    '``````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    For L = 0 To 19
        If Hash_Result(L) <> Decrypted_Data(L + 4 + Data_Length) Then Exit Function
    Next
    '``````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    ReDim Out_Decrypted(0 To Data_Length - 1)
    RtlMoveMemory Out_Decrypted(0), Decrypted_Data(4), Data_Length
    NextGen_DecryptData = True
    '`````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------'
End Function
'======================================================================================'


'=========================================================================================================================='
Private Function NextGen_Encrypt( _
                 ByRef Ptr_Data As LongPtr, _
                 ByRef Len_Data As Long, _
                 ByRef Input_IV As LongPtr, _
                 ByRef Input_Secret As LongPtr, _
                 ByRef Input_SecretLength As Long _
        ) As Byte()
'--------------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Handle_Alg  As LongPtr, Alg_ID As String
    Dim Command_DLL As String
    Dim Key_ObjectLength As Long
    Dim KeyObject_Byte() As Byte
    Dim IV_Length   As Long
    Dim IV_Bytes()  As Byte
    Dim Val         As String
    Dim Handle_Key  As LongPtr
    Dim Cipher_TextLength  As Long
    Dim CipherText_Bytes() As Byte
    Dim Data_Length As Long
    '```````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    Alg_ID = "AES" & vbNullChar
    BCryptOpenAlgorithmProvider Handle_Alg, StrPtr(Alg_ID), 0, 0
    '```````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    Command_DLL = "ObjectLength" & vbNullString
    BCryptGetProperty Handle_Alg, StrPtr(Command_DLL), Key_ObjectLength, LenB(Key_ObjectLength), 0, 0
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````'
    ReDim KeyObject_Byte(0 To Key_ObjectLength - 1): Command_DLL = "BlockLength" & vbNullChar
    BCryptGetProperty Handle_Alg, StrPtr(Command_DLL), IV_Length, LenB(IV_Length), 0, 0
    '````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````'
    ReDim IV_Bytes(0 To IV_Length - 1): RtlMoveMemory IV_Bytes(0), ByVal Input_IV, IV_Length
    '````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    Command_DLL = "ChainingMode" & vbNullString: Val = "ChainingModeCBC" & vbNullString
    BCryptSetProperty Handle_Alg, StrPtr(Command_DLL), ByVal StrPtr(Val), LenB(Val), 0&
    '``````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    BCryptGenerateSymmetricKey Handle_Alg, _
                               Handle_Key, KeyObject_Byte(0), Key_ObjectLength, ByVal Input_Secret, Input_SecretLength, 0
    BCryptEncrypt Handle_Key, ByVal Ptr_Data, Len_Data, ByVal 0, IV_Bytes(0), IV_Length, ByVal 0, 0, Cipher_TextLength, &H1
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    ReDim CipherText_Bytes(0 To Cipher_TextLength - 1)
    BCryptEncrypt Handle_Key, ByVal Ptr_Data, Len_Data, ByVal 0, IV_Bytes(0), IV_Length, _
                                  CipherText_Bytes(0), Cipher_TextLength, Data_Length, &H1
    '`````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    GoSub GS¦Clearing_Memory: NextGen_Encrypt = CipherText_Bytes: Exit Function
    '```````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Key) Then BCryptDestroyKey Handle_Key
    If CBool(Handle_Alg) Then BCryptCloseAlgorithmProvider Handle_Alg, 0&

    Return
'````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================================='


'======================================================================================================'
Private Function NextGen_Decrypt( _
                 ByRef Ptr_Data As LongPtr, _
                 ByRef Len_Data As Long, _
                 ByRef Ptr_IV As LongPtr, _
                 ByRef Ptr_Secret As LongPtr, _
                 ByRef Len_Secret As Long _
        ) As Byte()
'------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Handle_Alg  As LongPtr, Alg_ID As String
    Dim Command_DLL As String, Val     As String
    Dim Key_ObjectLength As Long
    Dim KeyObject_Byte() As Byte
    Dim IV_Length   As Long
    Dim IV_Bytes()  As Byte
    Dim Handle_Key  As LongPtr
    Dim Data_Length As Long
    Dim Output_Size As Long
    Dim Decrypted_Bytes() As Byte
    '```````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Alg_ID = "AES" & vbNullChar
    If BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr(Alg_ID), 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````'
    Command_DLL = "ObjectLength" & vbNullString
    If BCryptGetProperty(Handle_Alg, _
                         StrPtr(Command_DLL), Key_ObjectLength, LenB(Key_ObjectLength), 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````'
    ReDim KeyObject_Byte(0 To Key_ObjectLength - 1)
    Command_DLL = "BlockLength" & vbNullChar
    If BCryptGetProperty(Handle_Alg, StrPtr(Command_DLL), IV_Length, LenB(IV_Length), 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    ReDim IV_Bytes(0 To IV_Length - 1)
    RtlMoveMemory IV_Bytes(0), ByVal Ptr_IV, IV_Length
    '`````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Command_DLL = "ChainingMode" & vbNullString
    Val = "ChainingModeCBC" & vbNullString
    If BCryptSetProperty(Handle_Alg, _
                         StrPtr(Command_DLL), ByVal StrPtr(Val), LenB(Val), 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    If BCryptGenerateSymmetricKey(Handle_Alg, _
                                  Handle_Key, KeyObject_Byte(1), Key_ObjectLength, _
                                  ByVal Ptr_Secret, Len_Secret, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    If BCryptDecrypt(Handle_Key, _
                     ByVal Ptr_Data, Len_Data, ByVal 0, IV_Bytes(0), _
                     IV_Length, ByVal 0, 0, Output_Size, &H1) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    ReDim Decrypted_Bytes(0 To Output_Size - 1)
    If BCryptDecrypt(Handle_Key, _
                     ByVal Ptr_Data, Len_Data, ByVal 0, IV_Bytes(0), IV_Length, _
                     Decrypted_Bytes(0), Output_Size, Data_Length, &H1) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    NextGen_Decrypt = Decrypted_Bytes: Exit Function
    '```````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Key) Then BCryptDestroyKey Handle_Key
    If CBool(Handle_Alg) Then BCryptCloseAlgorithmProvider Handle_Alg, 0&

    Return
'````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'================================================================================================='
Private Function NextGen_HashGeneration_Wrapper( _
                 ByRef Data() As Byte, _
                 Optional ByRef Algorithm As String = "SHA1" _
        ) As Byte()
'-------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    NextGen_HashGeneration_Wrapper = _
    NextGen_HashGeneration(VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, Algorithm)
    '`````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='
'================================================================================================='
Private Function NextGen_HashGeneration( _
                ByRef Ptr_Data As LongPtr, _
                ByRef Len_Data As Long, _
                Optional ByRef Algorithm As String = "SHA1" _
        ) As Byte()
'-------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````'
    Dim Handle_Alg As LongPtr, Command_DLL As String
    Dim HashObject_Bytes() As Byte
    Dim Hash_Length  As Long, Length As Long
    Dim Hash_Bytes() As Byte, Alg_ID As String
    Dim Handle_Hash  As LongPtr
    '```````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    Alg_ID = Algorithm & vbNullChar
    If BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr(Alg_ID), 0, 0) Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    Command_DLL = "ObjectLength" & vbNullString
    If BCryptGetProperty(Handle_Alg, _
                         StrPtr(Command_DLL), Length, LenB(Length), 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    ReDim HashObject_Bytes(0 To Length - 1)
    Command_DLL = "HashDigestLength" & vbNullChar
    If BCryptGetProperty(Handle_Alg, _
                         StrPtr(Command_DLL), Hash_Length, LenB(Hash_Length), 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    ReDim Hash_Bytes(0 To Hash_Length - 1)
    If BCryptCreateHash(Handle_Alg, _
                        Handle_Hash, HashObject_Bytes(0), Length, 0, 0, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If BCryptHashData(Handle_Hash, ByVal Ptr_Data, Len_Data) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    If BCryptFinishHash(Handle_Hash, Hash_Bytes(0), Hash_Length, 0) <> 0 Then
        GoSub GS¦Clearing_Memory: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    NextGen_HashGeneration = Hash_Bytes: Exit Function
    '`````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Hash) Then BCryptDestroyHash Handle_Hash
    If CBool(Handle_Alg) Then BCryptCloseAlgorithmProvider Handle_Alg, 0

    Return
'```````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'=============================================================================='
Private Sub NextGen_Random_Wrapper( _
            ByRef Data() As Byte, _
            Optional ByRef Algorithm As String = "RNG" _
        )
'------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````'
    If LBound(Data) = -1 Then Exit Sub
    Call NextGen_Random(VarPtr(Data(LBound(Data))), _
                                    UBound(Data) - LBound(Data) + 1, Algorithm)
    '``````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------'
End Sub
'=============================================================================='
'======================================================================='
Private Sub NextGen_Random( _
            ByRef Ptr_Data As LongPtr, _
            ByRef Len_Data As Long, _
            Optional ByRef Algorithm As String = "RNG" _
        )
'-----------------------------------------------------------------------'

    '````````````````````````'
    Dim Handle_Alg As LongPtr
    Dim Alg_ID     As String
    '````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    Alg_ID = Algorithm & vbNullChar
    Call BCryptOpenAlgorithmProvider(Handle_Alg, StrPtr(Alg_ID), 0&, 0&)
    '```````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    Call BCryptGenRandom(Handle_Alg, ByVal Ptr_Data, Len_Data, 0&)
    Call BCryptCloseAlgorithmProvider(Handle_Alg, 0&)
    '`````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------'
End Sub
'======================================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'========================================================================================================'
Private Function AES_256_Encrypt_CryptoAPI( _
                 ByRef Initial_Data As String, _
                 ByRef Key_Crypt As String _
        ) As Byte()
'--------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````'
    Dim Handle_Prov  As LongPtr
    Dim Handle_Hash  As LongPtr
    Dim Handle_Key   As LongPtr
    Dim Data_Bytes() As Byte, Error_Message As String
    Dim Data_Len     As Long, Buffer_Len    As Long
    '````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    If CryptAcquireContext(Handle_Prov, _
                           vbNullString, MS_ENH_RSA_AES_PROV, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        Error_Message = "CryptDeriveKey: Ошибка получения криптографического контекста: " _
                                                                         & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Encrypt_CryptoAPI = Initial_Data: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````'
    If CryptCreateHash(Handle_Prov, CALG_SHA_256, 0, 0, Handle_Hash) = 0 Then
        Error_Message = "CryptDeriveKey: Ошибка создания хэша: " _
                                                & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Encrypt_CryptoAPI = Initial_Data: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Dim Password_Bytes() As Byte
    Password_Bytes = StrConv(Key_Crypt, vbFromUnicode)
    If CryptHashData(Handle_Hash, Password_Bytes(0), UBound(Password_Bytes) + 1, 0) = 0 Then
        Error_Message = "CryptDeriveKey: Ошибка хэширования пароля: " _
                                                     & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Encrypt_CryptoAPI = Initial_Data: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    If CryptDeriveKey(Handle_Prov, CALG_AES_256, Handle_Hash, 0, Handle_Key) = 0 Then
        Error_Message = "CryptDeriveKey: Ошибка создания ключа: " _
                                                 & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Encrypt_CryptoAPI = Initial_Data: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Data_Bytes = StrConv(Initial_Data, vbFromUnicode): Data_Len = UBound(Data_Bytes) + 1
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Buffer_Len = ((Data_Len \ 16) + 1) * 16: ReDim Preserve Data_Bytes(Buffer_Len - 1)
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    If CryptEncrypt(Handle_Key, 0, 1, 0, Data_Bytes(0), Data_Len, Buffer_Len) = 0 Then
        Error_Message = "CryptEncrypt: Ошибка шифрования: " _
                                           & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Encrypt_CryptoAPI = Initial_Data: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    AES_256_Encrypt_CryptoAPI = Data_Bytes: Exit Function
    '````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Key) Then CryptDestroyKey Handle_Key
    If CBool(Handle_Hash) Then CryptDestroyHash Handle_Hash
    If CBool(Handle_Prov) Then CryptReleaseContext Handle_Prov, 0&

    Return
'`````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


'========================================================================================================'
Private Function AES_256_Decrypt_CryptoAPI( _
                 ByRef Encrypted_Data As String, _
                 ByRef Key_Crypt As String _
        ) As String
'--------------------------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Handle_Prov   As LongPtr
    Dim Handle_Hash   As LongPtr
    Dim Handle_Key    As LongPtr
    Dim Data_Bytes()     As Byte
    Dim Data_Len         As Long
    Dim Password_Bytes() As Byte
    Dim Error_Message  As String
    '```````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    Data_Bytes = StrConv(Encrypted_Data, vbFromUnicode): Data_Len = UBound(Data_Bytes) + 1
    '`````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    If CryptAcquireContext(Handle_Prov, _
                           vbNullString, MS_ENH_RSA_AES_PROV, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        Error_Message = "CryptAcquireContext: Ошибка получения криптографического контекста: " _
                                                                              & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Decrypt_CryptoAPI = Encrypted_Data: Exit Function

    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````'
    If CryptCreateHash(Handle_Prov, CALG_SHA_256, 0, 0, Handle_Hash) = 0 Then
        Error_Message = "CryptCreateHash: Ошибка создания хэша: " & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Decrypt_CryptoAPI = Encrypted_Data: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Password_Bytes = StrConv(Key_Crypt, vbFromUnicode)

    If CryptHashData(Handle_Hash, _
                     Password_Bytes(0), UBound(Password_Bytes) + 1, 0) = 0 Then
        Error_Message = "CryptHashData: Ошибка хэширования пароля: " & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Decrypt_CryptoAPI = Encrypted_Data: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    If CryptDeriveKey(Handle_Prov, CALG_AES_256, Handle_Hash, 0, Handle_Key) = 0 Then
        Error_Message = "CryptDeriveKey: Ошибка создания ключа: " & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Decrypt_CryptoAPI = Encrypted_Data: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    If CryptDecrypt(Handle_Key, 0, 1, 0, Data_Bytes(0), Data_Len) = 0 Then
        Error_Message = "CryptDecrypt: Ошибка расшифровки: " & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: AES_256_Decrypt_CryptoAPI = Encrypted_Data: Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    ReDim Preserve Data_Bytes(Data_Len - 1): GoSub GS¦Clearing_Memory
    AES_256_Decrypt_CryptoAPI = StrConv(Data_Bytes, vbUnicode): Exit Function
    '````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Handle_Key) Then CryptDestroyKey Handle_Key
    If CBool(Handle_Hash) Then CryptDestroyHash Handle_Hash
    If CBool(Handle_Prov) Then CryptReleaseContext Handle_Prov, 0&

    Return
'`````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'==============================================================================='
Private Function XOR_Crypt_VBACode( _
                 ByRef Initial_Data As String, _
                 ByVal Is_Encode As Boolean, _
                 ByRef Key_Crypt As String _
        ) As String
'-------------------------------------------------------------------------------'

    '``````````````````````````'
    Dim Byte_Key()      As Byte
    Dim Byte_Swap(255)  As Byte
    Dim Vector_Result() As Byte
    Dim K As Long, M As Long
    Dim N As Long, i As Long
    '``````````````````````````'

    '```````````````````````````````````'
    Byte_Key = Initialize_Key(Key_Crypt)
    '```````````````````````````````````'

    '``````````````````````````````````````````'
    For i = 0 To 255: Byte_Swap(i) = i: Next i
    Call Mixing_Bytes(Byte_Key, Byte_Swap)
    '``````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    If Is_Encode Then
        If Len(Initial_Data) < 16 Then
            Initial_Data = Initial_Data & String$(8, vbNullChar)
        End If
    End If
    '```````````````````````````````````````````````````````````'

    '``````````````````````````````````````````'
    ReDim Vector_Result(1 To Len(Initial_Data))
    '``````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    For i = 1 To UBound(Vector_Result, 1)
        M = i Mod 256: K = (K + Byte_Swap(M)) Mod 256
        Call Swap_Bytes(Byte_Swap(M), Byte_Swap(K))
        N = Empty: N = N + Byte_Swap(K) + Byte_Swap(M)
        Vector_Result(i) = Asc(Mid$(Initial_Data, i, 1)) Xor Byte_Swap(N Mod 256)
    Next i
    '```````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    XOR_Crypt_VBACode = StrConv(Vector_Result, vbUnicode)
    '```````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


'======================================================================'
Private Function Initialize_Key( _
                 ByRef Key_Crypt As String _
        ) As Byte()
'----------------------------------------------------------------------'

    '`````````````````````````````'
    Dim Byte_Result(255)   As Byte
    Dim Len_Key As Long, i As Long
    '`````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Byte_Result(0) = Asc(Left$(Key_Crypt, 1)): Len_Key = Len(Key_Crypt)
    '``````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    For i = 1 To 255
        Byte_Result(i) = Asc(Mid$(Key_Crypt, 1 + (i Mod Len_Key), 1))
    Next i
    '``````````````````````````````````````````````````````````````````'

    '```````````````````````````'
    Initialize_Key = Byte_Result
    '```````````````````````````'

'----------------------------------------------------------------------'
End Function
'======================================================================'


'================================='
Private Sub Swap_Bytes( _
            ByRef B_1 As Byte, _
            ByRef B_2 As Byte _
        )
'---------------------------------'

    '```````````````'
    B_1 = B_1 Xor B_2
    B_2 = B_1 Xor B_2
    B_1 = B_1 Xor B_2
    '```````````````'

'---------------------------------'
End Sub
'================================='


'=========================================================='
Private Sub Mixing_Bytes( _
            ByRef Byte_Key() As Byte, _
            ByRef Byte_Swap() As Byte _
        )
'----------------------------------------------------------'

    '```````````````````````'
    Dim i As Long, K As Long
    '```````````````````````'

    '``````````````````````````````````````````````````````'
    For i = 0 To 255
        K = K + Byte_Key(i) + Byte_Swap(i)
        Call Swap_Bytes(Byte_Swap(i), Byte_Swap(K Mod 256))
    Next i
    '``````````````````````````````````````````````````````'

'----------------------------------------------------------'
End Sub
'=========================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'========================================================================================================='
Private Function AES_256_Crypt_VBACode( _
                 ByRef Initial_Data As String, _
                 ByRef Key_Crypt As String, _
                 ByVal Is_Encode As Boolean _
        ) As String
'---------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````'
    Dim sBox() As Variant
    Dim rCon() As Variant, sBoxInv() As Variant
    Dim g2()   As Variant, g3()      As Variant
    Dim g9()   As Variant, g11()     As Variant
    Dim g13()  As Variant, g14()     As Variant
    '```````````````````````````````````````````````````````````````````````````````````'
    Dim ExpandedKey As Variant, Block(16) As Variant, AESKey(32) As Variant
    Dim i As Long, IsDone As Boolean, J As Long
    Dim sPlain As Variant, sPass As Variant, sCipher As String, sTemp As Variant
    Dim Nonce(16) As Variant, PriorCipher(16) As Variant
    Dim x As Variant, R As Variant, Y As Variant, temp(4) As Variant, IntTemp As Variant
    '```````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    g2 = Array( _
        &H0, &H2, &H4, &H6, &H8, &HA, &HC, &HE, &H10, &H12, &H14, &H16, &H18, &H1A, &H1C, &H1E, _
        &H20, &H22, &H24, &H26, &H28, &H2A, &H2C, &H2E, &H30, &H32, &H34, &H36, &H38, &H3A, &H3C, &H3E, _
        &H40, &H42, &H44, &H46, &H48, &H4A, &H4C, &H4E, &H50, &H52, &H54, &H56, &H58, &H5A, &H5C, &H5E, _
        &H60, &H62, &H64, &H66, &H68, &H6A, &H6C, &H6E, &H70, &H72, &H74, &H76, &H78, &H7A, &H7C, &H7E, _
        &H80, &H82, &H84, &H86, &H88, &H8A, &H8C, &H8E, &H90, &H92, &H94, &H96, &H98, &H9A, &H9C, &H9E, _
        &HA0, &HA2, &HA4, &HA6, &HA8, &HAA, &HAC, &HAE, &HB0, &HB2, &HB4, &HB6, &HB8, &HBA, &HBC, &HBE, _
        &HC0, &HC2, &HC4, &HC6, &HC8, &HCA, &HCC, &HCE, &HD0, &HD2, &HD4, &HD6, &HD8, &HDA, &HDC, &HDE, _
        &HE0, &HE2, &HE4, &HE6, &HE8, &HEA, &HEC, &HEE, &HF0, &HF2, &HF4, &HF6, &HF8, &HFA, &HFC, &HFE, _
        &H1B, &H19, &H1F, &H1D, &H13, &H11, &H17, &H15, &HB, &H9, &HF, &HD, &H3, &H1, &H7, &H5, _
        &H3B, &H39, &H3F, &H3D, &H33, &H31, &H37, &H35, &H2B, &H29, &H2F, &H2D, &H23, &H21, &H27, &H25, _
        &H5B, &H59, &H5F, &H5D, &H53, &H51, &H57, &H55, &H4B, &H49, &H4F, &H4D, &H43, &H41, &H47, &H45, _
        &H7B, &H79, &H7F, &H7D, &H73, &H71, &H77, &H75, &H6B, &H69, &H6F, &H6D, &H63, &H61, &H67, &H65, _
        &H9B, &H99, &H9F, &H9D, &H93, &H91, &H97, &H95, &H8B, &H89, &H8F, &H8D, &H83, &H81, &H87, &H85, _
        &HBB, &HB9, &HBF, &HBD, &HB3, &HB1, &HB7, &HB5, &HAB, &HA9, &HAF, &HAD, &HA3, &HA1, &HA7, &HA5, _
        &HDB, &HD9, &HDF, &HDD, &HD3, &HD1, &HD7, &HD5, &HCB, &HC9, &HCF, &HCD, &HC3, &HC1, &HC7, &HC5, _
        &HFB, &HF9, &HFF, &HFD, &HF3, &HF1, &HF7, &HF5, &HEB, &HE9, &HEF, &HED, &HE3, &HE1, &HE7, &HE5)

    g3 = Array( _
        &H0, &H3, &H6, &H5, &HC, &HF, &HA, &H9, &H18, &H1B, &H1E, &H1D, &H14, &H17, &H12, &H11, _
        &H30, &H33, &H36, &H35, &H3C, &H3F, &H3A, &H39, &H28, &H2B, &H2E, &H2D, &H24, &H27, &H22, &H21, _
        &H60, &H63, &H66, &H65, &H6C, &H6F, &H6A, &H69, &H78, &H7B, &H7E, &H7D, &H74, &H77, &H72, &H71, _
        &H50, &H53, &H56, &H55, &H5C, &H5F, &H5A, &H59, &H48, &H4B, &H4E, &H4D, &H44, &H47, &H42, &H41, _
        &HC0, &HC3, &HC6, &HC5, &HCC, &HCF, &HCA, &HC9, &HD8, &HDB, &HDE, &HDD, &HD4, &HD7, &HD2, &HD1, _
        &HF0, &HF3, &HF6, &HF5, &HFC, &HFF, &HFA, &HF9, &HE8, &HEB, &HEE, &HED, &HE4, &HE7, &HE2, &HE1, _
        &HA0, &HA3, &HA6, &HA5, &HAC, &HAF, &HAA, &HA9, &HB8, &HBB, &HBE, &HBD, &HB4, &HB7, &HB2, &HB1, _
        &H90, &H93, &H96, &H95, &H9C, &H9F, &H9A, &H99, &H88, &H8B, &H8E, &H8D, &H84, &H87, &H82, &H81, _
        &H9B, &H98, &H9D, &H9E, &H97, &H94, &H91, &H92, &H83, &H80, &H85, &H86, &H8F, &H8C, &H89, &H8A, _
        &HAB, &HA8, &HAD, &HAE, &HA7, &HA4, &HA1, &HA2, &HB3, &HB0, &HB5, &HB6, &HBF, &HBC, &HB9, &HBA, _
        &HFB, &HF8, &HFD, &HFE, &HF7, &HF4, &HF1, &HF2, &HE3, &HE0, &HE5, &HE6, &HEF, &HEC, &HE9, &HEA, _
        &HCB, &HC8, &HCD, &HCE, &HC7, &HC4, &HC1, &HC2, &HD3, &HD0, &HD5, &HD6, &HDF, &HDC, &HD9, &HDA, _
        &H5B, &H58, &H5D, &H5E, &H57, &H54, &H51, &H52, &H43, &H40, &H45, &H46, &H4F, &H4C, &H49, &H4A, _
        &H6B, &H68, &H6D, &H6E, &H67, &H64, &H61, &H62, &H73, &H70, &H75, &H76, &H7F, &H7C, &H79, &H7A, _
        &H3B, &H38, &H3D, &H3E, &H37, &H34, &H31, &H32, &H23, &H20, &H25, &H26, &H2F, &H2C, &H29, &H2A, _
        &HB, &H8, &HD, &HE, &H7, &H4, &H1, &H2, &H13, &H10, &H15, &H16, &H1F, &H1C, &H19, &H1A)

    g9 = Array( _
        &H0, &H9, &H12, &H1B, &H24, &H2D, &H36, &H3F, &H48, &H41, &H5A, &H53, &H6C, &H65, &H7E, &H77, _
        &H90, &H99, &H82, &H8B, &HB4, &HBD, &HA6, &HAF, &HD8, &HD1, &HCA, &HC3, &HFC, &HF5, &HEE, &HE7, _
        &H3B, &H32, &H29, &H20, &H1F, &H16, &HD, &H4, &H73, &H7A, &H61, &H68, &H57, &H5E, &H45, &H4C, _
        &HAB, &HA2, &HB9, &HB0, &H8F, &H86, &H9D, &H94, &HE3, &HEA, &HF1, &HF8, &HC7, &HCE, &HD5, &HDC, _
        &H76, &H7F, &H64, &H6D, &H52, &H5B, &H40, &H49, &H3E, &H37, &H2C, &H25, &H1A, &H13, &H8, &H1, _
        &HE6, &HEF, &HF4, &HFD, &HC2, &HCB, &HD0, &HD9, &HAE, &HA7, &HBC, &HB5, &H8A, &H83, &H98, &H91, _
        &H4D, &H44, &H5F, &H56, &H69, &H60, &H7B, &H72, &H5, &HC, &H17, &H1E, &H21, &H28, &H33, &H3A, _
        &HDD, &HD4, &HCF, &HC6, &HF9, &HF0, &HEB, &HE2, &H95, &H9C, &H87, &H8E, &HB1, &HB8, &HA3, &HAA, _
        &HEC, &HE5, &HFE, &HF7, &HC8, &HC1, &HDA, &HD3, &HA4, &HAD, &HB6, &HBF, &H80, &H89, &H92, &H9B, _
        &H7C, &H75, &H6E, &H67, &H58, &H51, &H4A, &H43, &H34, &H3D, &H26, &H2F, &H10, &H19, &H2, &HB, _
        &HD7, &HDE, &HC5, &HCC, &HF3, &HFA, &HE1, &HE8, &H9F, &H96, &H8D, &H84, &HBB, &HB2, &HA9, &HA0, _
        &H47, &H4E, &H55, &H5C, &H63, &H6A, &H71, &H78, &HF, &H6, &H1D, &H14, &H2B, &H22, &H39, &H30, _
        &H9A, &H93, &H88, &H81, &HBE, &HB7, &HAC, &HA5, &HD2, &HDB, &HC0, &HC9, &HF6, &HFF, &HE4, &HED, _
        &HA, &H3, &H18, &H11, &H2E, &H27, &H3C, &H35, &H42, &H4B, &H50, &H59, &H66, &H6F, &H74, &H7D, _
        &HA1, &HA8, &HB3, &HBA, &H85, &H8C, &H97, &H9E, &HE9, &HE0, &HFB, &HF2, &HCD, &HC4, &HDF, &HD6, _
        &H31, &H38, &H23, &H2A, &H15, &H1C, &H7, &HE, &H79, &H70, &H6B, &H62, &H5D, &H54, &H4F, &H46)

    g11 = Array( _
        &H0, &HB, &H16, &H1D, &H2C, &H27, &H3A, &H31, &H58, &H53, &H4E, &H45, &H74, &H7F, &H62, &H69, _
        &HB0, &HBB, &HA6, &HAD, &H9C, &H97, &H8A, &H81, &HE8, &HE3, &HFE, &HF5, &HC4, &HCF, &HD2, &HD9, _
        &H7B, &H70, &H6D, &H66, &H57, &H5C, &H41, &H4A, &H23, &H28, &H35, &H3E, &HF, &H4, &H19, &H12, _
        &HCB, &HC0, &HDD, &HD6, &HE7, &HEC, &HF1, &HFA, &H93, &H98, &H85, &H8E, &HBF, &HB4, &HA9, &HA2, _
        &HF6, &HFD, &HE0, &HEB, &HDA, &HD1, &HCC, &HC7, &HAE, &HA5, &HB8, &HB3, &H82, &H89, &H94, &H9F, _
        &H46, &H4D, &H50, &H5B, &H6A, &H61, &H7C, &H77, &H1E, &H15, &H8, &H3, &H32, &H39, &H24, &H2F, _
        &H8D, &H86, &H9B, &H90, &HA1, &HAA, &HB7, &HBC, &HD5, &HDE, &HC3, &HC8, &HF9, &HF2, &HEF, &HE4, _
        &H3D, &H36, &H2B, &H20, &H11, &H1A, &H7, &HC, &H65, &H6E, &H73, &H78, &H49, &H42, &H5F, &H54, _
        &HF7, &HFC, &HE1, &HEA, &HDB, &HD0, &HCD, &HC6, &HAF, &HA4, &HB9, &HB2, &H83, &H88, &H95, &H9E, _
        &H47, &H4C, &H51, &H5A, &H6B, &H60, &H7D, &H76, &H1F, &H14, &H9, &H2, &H33, &H38, &H25, &H2E, _
        &H8C, &H87, &H9A, &H91, &HA0, &HAB, &HB6, &HBD, &HD4, &HDF, &HC2, &HC9, &HF8, &HF3, &HEE, &HE5, _
        &H3C, &H37, &H2A, &H21, &H10, &H1B, &H6, &HD, &H64, &H6F, &H72, &H79, &H48, &H43, &H5E, &H55, _
        &H1, &HA, &H17, &H1C, &H2D, &H26, &H3B, &H30, &H59, &H52, &H4F, &H44, &H75, &H7E, &H63, &H68, _
        &HB1, &HBA, &HA7, &HAC, &H9D, &H96, &H8B, &H80, &HE9, &HE2, &HFF, &HF4, &HC5, &HCE, &HD3, &HD8, _
        &H7A, &H71, &H6C, &H67, &H56, &H5D, &H40, &H4B, &H22, &H29, &H34, &H3F, &HE, &H5, &H18, &H13, _
        &HCA, &HC1, &HDC, &HD7, &HE6, &HED, &HF0, &HFB, &H92, &H99, &H84, &H8F, &HBE, &HB5, &HA8, &HA3)

    g13 = Array( _
        &H0, &HD, &H1A, &H17, &H34, &H39, &H2E, &H23, &H68, &H65, &H72, &H7F, &H5C, &H51, &H46, &H4B, _
        &HD0, &HDD, &HCA, &HC7, &HE4, &HE9, &HFE, &HF3, &HB8, &HB5, &HA2, &HAF, &H8C, &H81, &H96, &H9B, _
        &HBB, &HB6, &HA1, &HAC, &H8F, &H82, &H95, &H98, &HD3, &HDE, &HC9, &HC4, &HE7, &HEA, &HFD, &HF0, _
        &H6B, &H66, &H71, &H7C, &H5F, &H52, &H45, &H48, &H3, &HE, &H19, &H14, &H37, &H3A, &H2D, &H20, _
        &H6D, &H60, &H77, &H7A, &H59, &H54, &H43, &H4E, &H5, &H8, &H1F, &H12, &H31, &H3C, &H2B, &H26, _
        &HBD, &HB0, &HA7, &HAA, &H89, &H84, &H93, &H9E, &HD5, &HD8, &HCF, &HC2, &HE1, &HEC, &HFB, &HF6, _
        &HD6, &HDB, &HCC, &HC1, &HE2, &HEF, &HF8, &HF5, &HBE, &HB3, &HA4, &HA9, &H8A, &H87, &H90, &H9D, _
        &H6, &HB, &H1C, &H11, &H32, &H3F, &H28, &H25, &H6E, &H63, &H74, &H79, &H5A, &H57, &H40, &H4D, _
        &HDA, &HD7, &HC0, &HCD, &HEE, &HE3, &HF4, &HF9, &HB2, &HBF, &HA8, &HA5, &H86, &H8B, &H9C, &H91, _
        &HA, &H7, &H10, &H1D, &H3E, &H33, &H24, &H29, &H62, &H6F, &H78, &H75, &H56, &H5B, &H4C, &H41, _
        &H61, &H6C, &H7B, &H76, &H55, &H58, &H4F, &H42, &H9, &H4, &H13, &H1E, &H3D, &H30, &H27, &H2A, _
        &HB1, &HBC, &HAB, &HA6, &H85, &H88, &H9F, &H92, &HD9, &HD4, &HC3, &HCE, &HED, &HE0, &HF7, &HFA, _
        &HB7, &HBA, &HAD, &HA0, &H83, &H8E, &H99, &H94, &HDF, &HD2, &HC5, &HC8, &HEB, &HE6, &HF1, &HFC, _
        &H67, &H6A, &H7D, &H70, &H53, &H5E, &H49, &H44, &HF, &H2, &H15, &H18, &H3B, &H36, &H21, &H2C, _
        &HC, &H1, &H16, &H1B, &H38, &H35, &H22, &H2F, &H64, &H69, &H7E, &H73, &H50, &H5D, &H4A, &H47, _
        &HDC, &HD1, &HC6, &HCB, &HE8, &HE5, &HF2, &HFF, &HB4, &HB9, &HAE, &HA3, &H80, &H8D, &H9A, &H97)

    g14 = Array( _
        &H0, &HE, &H1C, &H12, &H38, &H36, &H24, &H2A, &H70, &H7E, &H6C, &H62, &H48, &H46, &H54, &H5A, _
        &HE0, &HEE, &HFC, &HF2, &HD8, &HD6, &HC4, &HCA, &H90, &H9E, &H8C, &H82, &HA8, &HA6, &HB4, &HBA, _
        &HDB, &HD5, &HC7, &HC9, &HE3, &HED, &HFF, &HF1, &HAB, &HA5, &HB7, &HB9, &H93, &H9D, &H8F, &H81, _
        &H3B, &H35, &H27, &H29, &H3, &HD, &H1F, &H11, &H4B, &H45, &H57, &H59, &H73, &H7D, &H6F, &H61, _
        &HAD, &HA3, &HB1, &HBF, &H95, &H9B, &H89, &H87, &HDD, &HD3, &HC1, &HCF, &HE5, &HEB, &HF9, &HF7, _
        &H4D, &H43, &H51, &H5F, &H75, &H7B, &H69, &H67, &H3D, &H33, &H21, &H2F, &H5, &HB, &H19, &H17, _
        &H76, &H78, &H6A, &H64, &H4E, &H40, &H52, &H5C, &H6, &H8, &H1A, &H14, &H3E, &H30, &H22, &H2C, _
        &H96, &H98, &H8A, &H84, &HAE, &HA0, &HB2, &HBC, &HE6, &HE8, &HFA, &HF4, &HDE, &HD0, &HC2, &HCC, _
        &H41, &H4F, &H5D, &H53, &H79, &H77, &H65, &H6B, &H31, &H3F, &H2D, &H23, &H9, &H7, &H15, &H1B, _
        &HA1, &HAF, &HBD, &HB3, &H99, &H97, &H85, &H8B, &HD1, &HDF, &HCD, &HC3, &HE9, &HE7, &HF5, &HFB, _
        &H9A, &H94, &H86, &H88, &HA2, &HAC, &HBE, &HB0, &HEA, &HE4, &HF6, &HF8, &HD2, &HDC, &HCE, &HC0, _
        &H7A, &H74, &H66, &H68, &H42, &H4C, &H5E, &H50, &HA, &H4, &H16, &H18, &H32, &H3C, &H2E, &H20, _
        &HEC, &HE2, &HF0, &HFE, &HD4, &HDA, &HC8, &HC6, &H9C, &H92, &H80, &H8E, &HA4, &HAA, &HB8, &HB6, _
        &HC, &H2, &H10, &H1E, &H34, &H3A, &H28, &H26, &H7C, &H72, &H60, &H6E, &H44, &H4A, &H58, &H56, _
        &H37, &H39, &H2B, &H25, &HF, &H1, &H13, &H1D, &H47, &H49, &H5B, &H55, &H7F, &H71, &H63, &H6D, _
        &HD7, &HD9, &HCB, &HC5, &HEF, &HE1, &HF3, &HFD, &HA7, &HA9, &HBB, &HB5, &H9F, &H91, &H83, &H8D)

    sBox = Array( _
        &H63, &H7C, &H77, &H7B, &HF2, &H6B, &H6F, &HC5, &H30, &H1, &H67, &H2B, &HFE, &HD7, &HAB, &H76, _
        &HCA, &H82, &HC9, &H7D, &HFA, &H59, &H47, &HF0, &HAD, &HD4, &HA2, &HAF, &H9C, &HA4, &H72, &HC0, _
        &HB7, &HFD, &H93, &H26, &H36, &H3F, &HF7, &HCC, &H34, &HA5, &HE5, &HF1, &H71, &HD8, &H31, &H15, _
        &H4, &HC7, &H23, &HC3, &H18, &H96, &H5, &H9A, &H7, &H12, &H80, &HE2, &HEB, &H27, &HB2, &H75, _
        &H9, &H83, &H2C, &H1A, &H1B, &H6E, &H5A, &HA0, &H52, &H3B, &HD6, &HB3, &H29, &HE3, &H2F, &H84, _
        &H53, &HD1, &H0, &HED, &H20, &HFC, &HB1, &H5B, &H6A, &HCB, &HBE, &H39, &H4A, &H4C, &H58, &HCF, _
        &HD0, &HEF, &HAA, &HFB, &H43, &H4D, &H33, &H85, &H45, &HF9, &H2, &H7F, &H50, &H3C, &H9F, &HA8, _
        &H51, &HA3, &H40, &H8F, &H92, &H9D, &H38, &HF5, &HBC, &HB6, &HDA, &H21, &H10, &HFF, &HF3, &HD2, _
        &HCD, &HC, &H13, &HEC, &H5F, &H97, &H44, &H17, &HC4, &HA7, &H7E, &H3D, &H64, &H5D, &H19, &H73, _
        &H60, &H81, &H4F, &HDC, &H22, &H2A, &H90, &H88, &H46, &HEE, &HB8, &H14, &HDE, &H5E, &HB, &HDB, _
        &HE0, &H32, &H3A, &HA, &H49, &H6, &H24, &H5C, &HC2, &HD3, &HAC, &H62, &H91, &H95, &HE4, &H79, _
        &HE7, &HC8, &H37, &H6D, &H8D, &HD5, &H4E, &HA9, &H6C, &H56, &HF4, &HEA, &H65, &H7A, &HAE, &H8, _
        &HBA, &H78, &H25, &H2E, &H1C, &HA6, &HB4, &HC6, &HE8, &HDD, &H74, &H1F, &H4B, &HBD, &H8B, &H8A, _
        &H70, &H3E, &HB5, &H66, &H48, &H3, &HF6, &HE, &H61, &H35, &H57, &HB9, &H86, &HC1, &H1D, &H9E, _
        &HE1, &HF8, &H98, &H11, &H69, &HD9, &H8E, &H94, &H9B, &H1E, &H87, &HE9, &HCE, &H55, &H28, &HDF, _
        &H8C, &HA1, &H89, &HD, &HBF, &HE6, &H42, &H68, &H41, &H99, &H2D, &HF, &HB0, &H54, &HBB, &H16)

    sBoxInv = Array( _
        &H52, &H9, &H6A, &HD5, &H30, &H36, &HA5, &H38, &HBF, &H40, &HA3, &H9E, &H81, &HF3, &HD7, &HFB, _
        &H7C, &HE3, &H39, &H82, &H9B, &H2F, &HFF, &H87, &H34, &H8E, &H43, &H44, &HC4, &HDE, &HE9, &HCB, _
        &H54, &H7B, &H94, &H32, &HA6, &HC2, &H23, &H3D, &HEE, &H4C, &H95, &HB, &H42, &HFA, &HC3, &H4E, _
        &H8, &H2E, &HA1, &H66, &H28, &HD9, &H24, &HB2, &H76, &H5B, &HA2, &H49, &H6D, &H8B, &HD1, &H25, _
        &H72, &HF8, &HF6, &H64, &H86, &H68, &H98, &H16, &HD4, &HA4, &H5C, &HCC, &H5D, &H65, &HB6, &H92, _
        &H6C, &H70, &H48, &H50, &HFD, &HED, &HB9, &HDA, &H5E, &H15, &H46, &H57, &HA7, &H8D, &H9D, &H84, _
        &H90, &HD8, &HAB, &H0, &H8C, &HBC, &HD3, &HA, &HF7, &HE4, &H58, &H5, &HB8, &HB3, &H45, &H6, _
        &HD0, &H2C, &H1E, &H8F, &HCA, &H3F, &HF, &H2, &HC1, &HAF, &HBD, &H3, &H1, &H13, &H8A, &H6B, _
        &H3A, &H91, &H11, &H41, &H4F, &H67, &HDC, &HEA, &H97, &HF2, &HCF, &HCE, &HF0, &HB4, &HE6, &H73, _
        &H96, &HAC, &H74, &H22, &HE7, &HAD, &H35, &H85, &HE2, &HF9, &H37, &HE8, &H1C, &H75, &HDF, &H6E, _
        &H47, &HF1, &H1A, &H71, &H1D, &H29, &HC5, &H89, &H6F, &HB7, &H62, &HE, &HAA, &H18, &HBE, &H1B, _
        &HFC, &H56, &H3E, &H4B, &HC6, &HD2, &H79, &H20, &H9A, &HDB, &HC0, &HFE, &H78, &HCD, &H5A, &HF4, _
        &H1F, &HDD, &HA8, &H33, &H88, &H7, &HC7, &H31, &HB1, &H12, &H10, &H59, &H27, &H80, &HEC, &H5F, _
        &H60, &H51, &H7F, &HA9, &H19, &HB5, &H4A, &HD, &H2D, &HE5, &H7A, &H9F, &H93, &HC9, &H9C, &HEF, _
        &HA0, &HE0, &H3B, &H4D, &HAE, &H2A, &HF5, &HB0, &HC8, &HEB, &HBB, &H3C, &H83, &H53, &H99, &H61, _
        &H17, &H2B, &H4, &H7E, &HBA, &H77, &HD6, &H26, &HE1, &H69, &H14, &H63, &H55, &H21, &HC, &H7D)

    rCon = Array( _
        &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, _
        &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, _
        &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, _
        &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, _
        &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, _
        &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, _
        &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, _
        &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, _
        &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, _
        &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, _
        &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, _
        &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, _
        &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, _
        &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, _
        &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, _
        &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB)
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    For i = 0 To 15
        Nonce(i) = 0
    Next i

    For i = 0 To (Len(Key_Crypt) - 1)
        AESKey(i Mod 33) = Asc(Mid(Key_Crypt, i + 1, 1))
    Next i

    For i = Len(Key_Crypt) To 31
        AESKey(i) = 0
    Next i
    '```````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    ExpandedKey = Expand_Key(AESKey, sBox, rCon)
    '```````````````````````````````````````````'

    '`````````````````````'
    sPlain = Initial_Data
    sCipher = "": J = 0
    IsDone = False
    '`````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    Do Until IsDone

        sTemp = Mid(sPlain, J * 16 + 1, 16)

        If Len(sTemp) < 16 Then
            For i = Len(sTemp) To 15
                sTemp = sTemp & Chr(0)
            Next
        End If

        For i = 0 To 15
            Block(i) = Asc(Mid(sTemp, (i Mod 4) * 4 + (i \ 4) + 1, 1))
        Next

        If (J + 1) * 16 >= Len(sPlain) Then
            IsDone = True
        End If

        J = J + 1

        If Is_Encode Then
            R = 0
            For i = 0 To 15
                Block(i) = Block(i) Xor Nonce(i) Xor ExpandedKey((i Mod 4) * 4 + (i \ 4))
            Next

            For x = 1 To 13
                Block(0) = sBox(Block(0))
                Block(1) = sBox(Block(1))
                Block(2) = sBox(Block(2))
                Block(3) = sBox(Block(3))

                IntTemp = sBox(Block(4))
                Block(4) = sBox(Block(5))
                Block(5) = sBox(Block(6))
                Block(6) = sBox(Block(7))
                Block(7) = IntTemp

                IntTemp = sBox(Block(8))
                Block(8) = sBox(Block(10))
                Block(10) = IntTemp
                IntTemp = sBox(Block(9))
                Block(9) = sBox(Block(11))
                Block(11) = IntTemp

                IntTemp = sBox(Block(12))
                Block(12) = sBox(Block(15))
                Block(15) = sBox(Block(14))
                Block(14) = sBox(Block(13))
                Block(13) = IntTemp

                R = x * 16
                For i = 0 To 3
                    temp(0) = Block(i)
                    temp(1) = Block(i + 4)
                    temp(2) = Block(i + 8)
                    temp(3) = Block(i + 12)

                    Block(i) = g2(temp(0)) Xor temp(3) Xor temp(2) _
                                           Xor g3(temp(1)) Xor ExpandedKey(R + i * 4)
                    Block(i + 4) = g2(temp(1)) Xor temp(0) Xor temp(3) _
                                               Xor g3(temp(2)) Xor ExpandedKey(R + i * 4 + 1)
                    Block(i + 8) = g2(temp(2)) Xor temp(1) Xor temp(0) _
                                               Xor g3(temp(3)) Xor ExpandedKey(R + i * 4 + 2)
                    Block(i + 12) = g2(temp(3)) Xor temp(2) Xor temp(1) _
                                                Xor g3(temp(0)) Xor ExpandedKey(R + i * 4 + 3)
                Next
            Next

            Block(0) = sBox(Block(0)) Xor ExpandedKey(224)
            Block(1) = sBox(Block(1)) Xor ExpandedKey(228)
            Block(2) = sBox(Block(2)) Xor ExpandedKey(232)
            Block(3) = sBox(Block(3)) Xor ExpandedKey(236)

            IntTemp = sBox(Block(4)) Xor ExpandedKey(237)
            Block(4) = sBox(Block(5)) Xor ExpandedKey(225)
            Block(5) = sBox(Block(6)) Xor ExpandedKey(229)
            Block(6) = sBox(Block(7)) Xor ExpandedKey(233)
            Block(7) = IntTemp

            IntTemp = sBox(Block(8)) Xor ExpandedKey(234)
            Block(8) = sBox(Block(10)) Xor ExpandedKey(226)
            Block(10) = IntTemp
            IntTemp = sBox(Block(9)) Xor ExpandedKey(238)
            Block(9) = sBox(Block(11)) Xor ExpandedKey(230)
            Block(11) = IntTemp

            IntTemp = sBox(Block(12)) Xor ExpandedKey(231)
            Block(12) = sBox(Block(15)) Xor ExpandedKey(227)
            Block(15) = sBox(Block(14)) Xor ExpandedKey(239)
            Block(14) = sBox(Block(13)) Xor ExpandedKey(235)
            Block(13) = IntTemp

            For i = 0 To 15
                Nonce(i) = Block(i)
            Next
        Else
            For i = 0 To 15
                PriorCipher(i) = Block(i)
            Next

            Block(0) = sBoxInv(Block(0) Xor ExpandedKey(224))
            Block(1) = sBoxInv(Block(1) Xor ExpandedKey(228))
            Block(2) = sBoxInv(Block(2) Xor ExpandedKey(232))
            Block(3) = sBoxInv(Block(3) Xor ExpandedKey(236))

            IntTemp = sBoxInv(Block(4) Xor ExpandedKey(225))
            Block(4) = sBoxInv(Block(7) Xor ExpandedKey(237))
            Block(7) = sBoxInv(Block(6) Xor ExpandedKey(233))
            Block(6) = sBoxInv(Block(5) Xor ExpandedKey(229))
            Block(5) = IntTemp

            IntTemp = sBoxInv(Block(8) Xor ExpandedKey(226))
            Block(8) = sBoxInv(Block(10) Xor ExpandedKey(234))
            Block(10) = IntTemp
            IntTemp = sBoxInv(Block(9) Xor ExpandedKey(230))
            Block(9) = sBoxInv(Block(11) Xor ExpandedKey(238))
            Block(11) = IntTemp

            IntTemp = sBoxInv(Block(12) Xor ExpandedKey(227))
            Block(12) = sBoxInv(Block(13) Xor ExpandedKey(231))
            Block(13) = sBoxInv(Block(14) Xor ExpandedKey(235))
            Block(14) = sBoxInv(Block(15) Xor ExpandedKey(239))
            Block(15) = IntTemp

            For x = 13 To 1 Step -1
                R = x * 16

                For i = 0 To 3
                    temp(0) = Block(i) Xor ExpandedKey(R + i * 4)
                    temp(1) = Block(i + 4) Xor ExpandedKey(R + i * 4 + 1)
                    temp(2) = Block(i + 8) Xor ExpandedKey(R + i * 4 + 2)
                    temp(3) = Block(i + 12) Xor ExpandedKey(R + i * 4 + 3)

                    Block(i) = g14(temp(0)) Xor g9(temp(3)) Xor g13(temp(2)) Xor g11(temp(1))
                    Block(i + 4) = g14(temp(1)) Xor g9(temp(0)) Xor g13(temp(3)) Xor g11(temp(2))
                    Block(i + 8) = g14(temp(2)) Xor g9(temp(1)) Xor g13(temp(0)) Xor g11(temp(3))
                    Block(i + 12) = g14(temp(3)) Xor g9(temp(2)) Xor g13(temp(1)) Xor g11(temp(0))
                Next

                Block(0) = sBoxInv(Block(0))
                Block(1) = sBoxInv(Block(1))
                Block(2) = sBoxInv(Block(2))
                Block(3) = sBoxInv(Block(3))

                IntTemp = sBoxInv(Block(4))
                Block(4) = sBoxInv(Block(7))
                Block(7) = sBoxInv(Block(6))
                Block(6) = sBoxInv(Block(5))
                Block(5) = IntTemp

                IntTemp = sBoxInv(Block(8))
                Block(8) = sBoxInv(Block(10))
                Block(10) = IntTemp
                IntTemp = sBoxInv(Block(9))
                Block(9) = sBoxInv(Block(11))
                Block(11) = IntTemp

                IntTemp = sBoxInv(Block(12))
                Block(12) = sBoxInv(Block(13))
                Block(13) = sBoxInv(Block(14))
                Block(14) = sBoxInv(Block(15))
                Block(15) = IntTemp
            Next

            R = 0
            For i = 0 To 15
                Block(i) = Block(i) Xor ExpandedKey((i Mod 4) * 4 + (i \ 4)) Xor Nonce(i)
                Nonce(i) = PriorCipher(i)
            Next
        End If

        For i = 0 To 15
            sCipher = sCipher & Chr$(Block((i Mod 4) * 4 + (i \ 4)))
        Next i

    Loop
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````'
    AES_256_Crypt_VBACode = sCipher
    '``````````````````````````````'

'---------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================='


'==============================================================='
Private Function Expand_Key( _
                 ByRef Key() As Variant, _
                 ByRef sBox() As Variant, _
                 ByRef rCon() As Variant _
        ) As Variant()
'---------------------------------------------------------------'

    '```````````````````````````````````````'
    Dim rConIter As Variant
    Dim temp()   As Variant
    Dim i As Variant, result(240) As Variant
    '```````````````````````````````````````'

    '``````````````````````````'
    ReDim temp(4): rConIter = 1
    '``````````````````````````'

    '`````````````````````'
    For i = 0 To 31
        result(i) = Key(i)
    Next
    '`````````````````````'

    '```````````````````````````````````````````````````````````'
    For i = 32 To 239 Step 4
        temp(0) = result(i - 4)
        temp(1) = result(i - 3)
        temp(2) = result(i - 2)
        temp(3) = result(i - 1)

        If i Mod 32 = 0 Then
            temp = ScheduleCore_Key(temp, rConIter, sBox, rCon)
            rConIter = rConIter + 1
        End If

        If i Mod 32 = 16 Then
            temp(0) = sBox(temp(0))
            temp(1) = sBox(temp(1))
            temp(2) = sBox(temp(2))
            temp(3) = sBox(temp(3))
        End If

        result(i) = result(i - 32) Xor temp(0)
        result(i + 1) = result(i - 31) Xor temp(1)
        result(i + 2) = result(i - 30) Xor temp(2)
        result(i + 3) = result(i - 29) Xor temp(3)
    Next
    '```````````````````````````````````````````````````````````'

    '``````````````````'
    Expand_Key = result
    '``````````````````'

'---------------------------------------------------------------'
End Function
'==============================================================='


'==============================================================='
Private Function ScheduleCore_Key( _
                 ByRef Row() As Variant, _
                 ByVal A As Variant, _
                 ByRef sBox() As Variant, _
                 ByRef rCon() As Variant _
        ) As Variant()
'---------------------------------------------------------------'

    '``````````````````````````````````'
    Dim result(4) As Variant, i As Long
    '``````````````````````````````````'

    '```````````````````````````````````````'
    For i = 0 To 3
        result(i) = sBox(Row((i + 5) Mod 4))
    Next
    '```````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    result(0) = result(0) Xor rCon(A): ScheduleCore_Key = result
    '```````````````````````````````````````````````````````````'

'---------------------------------------------------------------'
End Function
'==============================================================='


'============================================================'
Private Function NullByte_Trim( _
                 ByRef Initial_Data As String _
        ) As String
'------------------------------------------------------------'

    '``````````````````````````````'
    Dim Position As Long, i As Long
    '``````````````````````````````'

    '````````````````````````````````````````````````````````'
    For i = Len(Initial_Data) To 1 Step -1
        Position = i
        If Mid$(Initial_Data, i, 1) <> ChrW$(0) Then Exit For
    Next i
    '````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    If Position = 1 Then
        If Mid$(Initial_Data, i, 1) = ChrW$(0) Then
            Position = Len(Initial_Data)
        End If
    End If
    '```````````````````````````````````````````````'

    '````````````````````````````````````````````'
    NullByte_Trim = Left$(Initial_Data, Position)
    '````````````````````````````````````````````'

'------------------------------------------------------------'
End Function
'============================================================'


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'============================================================='
Private Function Conversion_BytesToString( _
                 ByRef Array_Bytes() As Byte _
        ) As String
'-------------------------------------------------------------'
    Conversion_BytesToString = StrConv(Array_Bytes, vbUnicode)
'-------------------------------------------------------------'
End Function
'============================================================='


'============================================================='
Private Function Conversion_StringToBytes( _
                 ByVal Str_Data As String _
        ) As Byte()
'-------------------------------------------------------------'
    Conversion_StringToBytes = StrConv(Str_Data, vbFromUnicode)
'-------------------------------------------------------------'
End Function
'============================================================='


'======================================================================='
Private Function Base64_Encode_CryptoNG( _
                 ByRef Array_Data() As Byte _
        ) As String
'-----------------------------------------------------------------------'

    '```````````````````````````````````'
    Dim Output As String, Length As Long
    '```````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If CryptBinaryToString(Array_Data(0), UBound(Array_Data) + 1, _
       CRYPT_STRING_BASE64, vbNullString, Length) = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````````'

    '`````````````````````````````````'
    Output = String(Length - 1, Chr(0))
    '`````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    If CryptBinaryToString(Array_Data(0), UBound(Array_Data) + 1, _
       CRYPT_STRING_BASE64, Output, Length) = 0 Then Exit Function
    '``````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Base64_Encode_CryptoNG = Output
    '```````````````````````````````'

'-----------------------------------------------------------------------'
End Function
'======================================================================='


'========================================================================'
Private Function Base64_Decode_CryptoNG( _
                 ByVal Str_Data As String _
        ) As String
'------------------------------------------------------------------------'

    '```````````````````````````````````````'
    Dim Array_Data() As Byte, Length As Long
    '```````````````````````````````````````'

    '```````````````````````````````````````````````````````````````'
    If CryptStringToBinary(Str_Data, Len(Str_Data), _
       CRYPT_STRING_BASE64, ByVal 0&, Length) = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    ReDim Array_Data(0 To Length - 1)
    '```````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    If CryptStringToBinary(Str_Data, Len(Str_Data), _
       CRYPT_STRING_BASE64, Array_Data(0), Length) = 0 Then Exit Function
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````'
    Base64_Decode_CryptoNG = Array_Data
    '``````````````````````````````````'

'------------------------------------------------------------------------'
End Function
'========================================================================'


'==================================================='
Private Function Base64_Encode_Reserve( _
                 ByRef Array_Data() As Byte _
        ) As String
'---------------------------------------------------'

    '````````````````````````````````````````'
    Dim Obj_XML As Object, Obj_Node As Object
    '````````````````````````````````````````'

    '```````````````````````````````````````````````'
    Set Obj_XML = CreateObject("MSXML2.DOMDocument")
    Set Obj_Node = Obj_XML.createElement("b64")
    '```````````````````````````````````````````````'

    '````````````````````````````````````'
    Obj_Node.DataType = "bin.base64"
    Obj_Node.nodeTypedValue = Array_Data
    Base64_Encode_Reserve = Obj_Node.Text
    '````````````````````````````````````'

    '````````````````````````````````````````````'
    Set Obj_Node = Nothing: Set Obj_XML = Nothing
    '````````````````````````````````````````````'

'---------------------------------------------------'
End Function
'==================================================='


'==================================================='
Private Function Base64_Decode_Reserve( _
                 ByVal Str_Data As String _
        ) As Byte()
'---------------------------------------------------'

    '````````````````````````````````````````'
    Dim Obj_XML As Object, Obj_Node As Object
    '````````````````````````````````````````'

    '``````````````````````````````````````````````'
    Set Obj_XML = CreateObject("MSXML2.DOMDocument")
    Set Obj_Node = Obj_XML.createElement("b64")
    '``````````````````````````````````````````````'

    '``````````````````````````````````````````````'
    Obj_Node.DataType = "bin.base64"
    Obj_Node.Text = Str_Data
    Base64_Decode_Reserve = Obj_Node.nodeTypedValue
    '``````````````````````````````````````````````'

    '````````````````````````````````````````````'
    Set Obj_Node = Nothing: Set Obj_XML = Nothing
    '````````````````````````````````````````````'

'---------------------------------------------------'
End Function
'==================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================'
Private Function Get_SemanticDataType( _
                 ByRef Data As Variant _
        ) As MSOffice_Type_ContentFormat_Crypt
'------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````'
    Dim Data_Type      As FileSystem_Directory_Type, Error_Message As String
    Dim Data_VarType   As Long, Data_TypeName        As String
    Dim Obj_VBProjects As Object, Obj_VBComponent    As Object
    Dim VBProject_Name As String, VBComponent_Name   As String
    '```````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Data_VarType = VarType(Data): Data_TypeName = TypeName(Data)
    '``````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_VarType
        Case vbBoolean:              Get_SemanticDataType = MSDI_Bool_Crypt:        Exit Function
        Case vbEmpty:                Get_SemanticDataType = MSDI_Empty_Crypt:       Exit Function
        Case vbNull:                 Get_SemanticDataType = MSDI_Null_Crypt:        Exit Function
        Case vbError, vbObjectError: Get_SemanticDataType = MSDI_Undefined_Crypt:   Exit Function
        Case vbUserDefinedType:      Get_SemanticDataType = MSDI_UserDefType_Crypt: Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If (Data_VarType And vbArray) And (Not Data_TypeName = "Range") Then
        Get_SemanticDataType = MSDI_Array_Crypt: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_TypeName
        Case "String"
            If Len(Data) = 0 Then Get_SemanticDataType = MSDI_NullString_Crypt: Exit Function
            If IsNumeric(Data) Then Get_SemanticDataType = MSDI_String_Crypt:   Exit Function
            Data_Type = Get_Directory_Type(CStr(Data))

            Select Case Data_Type
                Case DirType_File:     Get_SemanticDataType = MSDI_File_Crypt:                  Exit Function
                Case DirType_Folder:   Get_SemanticDataType = MSDI_Folder_Crypt:                Exit Function
                Case DirType_NotFound: Get_SemanticDataType = MSDI_NonExistent_Directory_Crypt: Exit Function
                Case DirType_Invalid:
                    On Error Resume Next
                    VBProject_Name = Application.VBE.ActiveVBProject.Name

                    If Not Err.Number = 0 Then
                        Error_Message = "Невозможно проверить тип данных (Компонент это или VB-Проект)! " & _
                                        "Тип данных будет определен как строка (String)!"
                        Call Show_ErrorMessage_Immediate(Error_Message, "Проблема идентификации типов")
                        Get_SemanticDataType = MSDI_String_Crypt: On Error GoTo 0: Exit Function
                    End If

                    If VBProject_Name = Data Then
                        Get_SemanticDataType = MSDI_VBProject_Crypt
                        On Error GoTo 0: Exit Function
                    End If

                    Err.Clear
                    VBComponent_Name = Application.VBE.ActiveVBProject.VBComponents.Item(Data).Name
                    
                    If Len(VBComponent_Name) > 0 And Err.Number = 0 Then
                        Get_SemanticDataType = MSDI_VBComponent_Crypt
                        On Error GoTo 0: Exit Function
                    End If

                    For Each Obj_VBProjects In Application.VBE.VBProjects
                        If Obj_VBProjects.Name = Data Then
                            Get_SemanticDataType = MSDI_VBProject_Crypt:   On Error GoTo 0: Exit Function
                        End If

                        Err.Clear: Set Obj_VBComponent = Obj_VBProjects.VBComponents.Item(Data)
                        If (Err.Number = 0) And (Not Obj_VBComponent Is Nothing) Then
                            If Not Len(Obj_VBComponent.Name) = 0 Then
                                Set Obj_VBComponent = Nothing
                                Get_SemanticDataType = MSDI_VBComponent_Crypt: On Error GoTo 0: Exit Function
                            End If
                        End If
                        Set Obj_VBComponent = Nothing
                    Next Obj_VBProjects

                    Get_SemanticDataType = MSDI_String_Crypt: On Error GoTo 0: Exit Function
            End Select

        Case "Byte":       Get_SemanticDataType = MSDI_Int8_Crypt:             Exit Function
        Case "Integer":    Get_SemanticDataType = MSDI_Int16_Crypt:            Exit Function
        Case "Long":       Get_SemanticDataType = MSDI_Int32_Crypt:            Exit Function
        Case "LongLong":   Get_SemanticDataType = MSDI_Int64_Crypt:            Exit Function
        Case "Single":     Get_SemanticDataType = MSDI_FloatPoint32_Crypt:     Exit Function
        Case "Currency":   Get_SemanticDataType = MSDI_FloatPoint64Cur_Crypt:  Exit Function
        Case "Double":     Get_SemanticDataType = MSDI_FloatPoint64Dbl_Crypt:  Exit Function
        Case "Decimal":    Get_SemanticDataType = MSDI_FloatPoint112_Crypt:    Exit Function
        Case "Date":       Get_SemanticDataType = MSDI_Date_Crypt:             Exit Function
        Case "Range":      Get_SemanticDataType = MSDI_Range_Crypt:            Exit Function
        Case "Dictionary": Get_SemanticDataType = MSDI_Dictionary_Crypt:       Exit Function
        Case "Collection": Get_SemanticDataType = MSDI_Collection_Crypt:       Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````'
    If IsObject(Data) Then
        If Not Data Is Nothing Then
            Get_SemanticDataType = MSDI_Object_Crypt:  Exit Function
        Else
            Get_SemanticDataType = MSDI_Nothing_Crypt: Exit Function
        End If
    Else
        Get_SemanticDataType = MSDI_Undefined_Crypt:   Exit Function
    End If
    '```````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'==============================================================================================================='
Private Function Conversion_FileToBytes( _
                 ByRef File_Path As String _
        ) As Variant
'---------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````'
    Dim Return_HashSumm As String
    Dim Size_File As Double, FF As Integer
    Dim Return_Bytes() As Byte, Data_Bytes()  As Byte
    Dim Handle_File As Long, Flag_Done     As Boolean
    Dim Buffer_Bytes() As Byte, Handled_Bytes As Long
    '````````````````````````````````````````````````'
    Dim Total_HashSumm As String, Bytes_Read As Long
    Dim Read_StartPosition As Long, Len_Hash As Long
    Dim Error_Message As String
    '````````````````````````````````````````````````'
    #If Win64 Then
        Const BlockSize_Hashing As Long = &H1000000
    #Else
        Const BlockSize_Hashing As Long = &H3FFFFF
    #End If
    '``````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Conversion_FileToBytes = Empty
    
    If File_IsOpen(File_Path) Then
        Error_Message = "Файл открыт и не может быть прочитан (" & File_Path & ")!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу")
        Exit Function
    End If
    
    Size_File = Get_FileSize(File_Path)
    
    Select Case Size_File
        Case 0
            Error_Message = "Нет данных для хеширования! Размер файла равен 0 (" & File_Path & ")!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Отсутствуют данные для обработки")
            Exit Function
        Case -1
            Error_Message = "Не удалось открыть файл. Возможно проблема с правами доступа (" & File_Path & ")!"
            Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу")
            Exit Function
    End Select

    Select Case Size_File
        Case Is > 2147483648#
            Error_Message = "Невозможно прочитать файл (Размер файла более 2 Гб)"
            Call Show_ErrorMessage_Immediate(Error_Message, "Поддержка файлов более 2 Гб не реализованна!")
        Case Is > 536870912#
            #If Win64 Then
                GoSub GS¦ReadingFile_Less_2GB
            #Else
                GoSub GS¦Throw_OutOfMemory
            #End If
        Case Else
            GoSub GS¦ReadingFile_Less_2GB
    End Select

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````'
GS¦ReadingFile_Less_2GB:

    FF = FreeFile: Size_File = Size_File - 1
    
    On Error Resume Next
    ReDim Return_Bytes(0 To Size_File) As Byte

    If Err.Number <> 0 Then
        #If Not Win64 Then
            GoSub GS¦Throw_OutOfMemory
        #End If
        
        On Error GoTo 0: Return
    End If

    On Error GoTo 0

    Open File_Path For Binary Access Read As #FF Len = 1
    Get #FF, , Return_Bytes
    Close #FF
    
    On Error Resume Next
    Conversion_FileToBytes = Return_Bytes

    If Err.Number <> 0 Then
        #If Not Win64 Then
            GoSub GS¦Throw_OutOfMemory
        #End If
        
        Conversion_FileToBytes = Empty
    End If
    
    On Error GoTo 0

    Return
'```````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Throw_OutOfMemory:

    Error_Message = "Обновите MS Office до x64 для расширения максимального объёма памяти!"
    Call Show_ErrorMessage_Immediate(Error_Message, "Недостаточно виртуальной памяти для считывания файла")

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------------'
End Function
'==============================================================================================================='


'==============================================='
Private Function Conversion_BytesToFile( _
                 ByRef Data_Bytes() As Byte, _
                 ByRef File_Path As String _
        ) As Boolean
'-----------------------------------------------'

    '````````````````'
    Dim FF As Integer
    '````````````````'

    '```````````````````````````````````````````'
    FF = FreeFile(): On Error Resume Next
    Open File_Path For Binary Access Write As #FF
    Put #FF, , Data_Bytes: Close #FF
    '```````````````````````````````````````````'

    '``````````````````````````````````````'
    Conversion_BytesToFile = Err.Number = 0
    On Error GoTo 0
    '``````````````````````````````````````'

'-----------------------------------------------'
End Function
'==============================================='


'============================================'
Private Function File_IsOpen( _
                 ByRef File_Path As String _
        ) As Boolean
'--------------------------------------------'

    '``````````````````````'
    Dim File_Num As Integer
    '``````````````````````'

    '````````````````````````````````````````'
    On Error Resume Next: File_Num = FreeFile

    Open File_Path For Random Access _
    Read Write Lock Read Write As #File_Num
    Close #File_Num

    File_IsOpen = Err.Number: On Error GoTo 0
    '````````````````````````````````````````'

'--------------------------------------------'
End Function
'============================================'


'============================================================================'
Private Function Get_FileSize( _
                 ByVal FileSystem_FilePath As String, _
                 Optional Format_Size As Format_SizeFile = SizeFormat_Byte _
        ) As Double
'----------------------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Handle_File As LongPtr, result As Double
    Dim File_Size   As Large_Integer
    '```````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Handle_File = CreateFile( _
                  FileSystem_FilePath, GENERIC_READ, FILE_SHARE_READ, _
                  0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0& _
            )
    '``````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    If Handle_File = -1 Then Get_FileSize = -1: Exit Function
    '```````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    If GetFileSizeEx(Handle_File, File_Size) <> 0 Then
        result = File_Size.High_Part * 4294967296# + _
                 IIf(File_Size.Low_Part < 0, _
                     File_Size.Low_Part + 4294967296#, _
                     File_Size.Low_Part)

        Select Case Format_Size
            Case SizeFormat_KByte: result = result / 1024
            Case SizeFormat_MByte: result = result / 1048576
            Case SizeFormat_GByte: result = result / 1073741824
        End Select

        Get_FileSize = result
    Else
        Get_FileSize = -1
    End If

    CloseHandle Handle_File
    '```````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------'
End Function
'============================================================================'


'====================================================================='
Private Function Get_Directory_Type( _
                 ByRef FileSystem_Path As String _
        ) As FileSystem_Directory_Type
'---------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````'
    Dim File_Arrt      As VbFileAttribute, IsAbsolute_Path As Boolean
    Static Obj_RegExp  As Object, Obj_FSO  As Object
    '`````````````````````````````````````````````````````````````````'
    Const vb_Directory As Long = &H10
    '`````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If Obj_RegExp Is Nothing Then
        Set Obj_RegExp = CreateObject("VBScript.RegExp")
        Obj_RegExp.Pattern = _
        "^(?:[a-zA-Z]:|\\\\[\w.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])*"
    End If

    If Not Obj_RegExp.Test(FileSystem_Path) Then
        Get_Directory_Type = DirType_Invalid
    End If
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    If Get_Directory_Type = DirType_Invalid Then
        GoSub GS¦Absolute_Path
        If Not IsAbsolute_Path Then Exit Function
    Else
        IsAbsolute_Path = True
    End If
    '````````````````````````````````````````````'

    '``````````````````````````````````````````````````'
    On Error GoTo GT¦Directory_NotFound

    File_Arrt = GetAttr(FileSystem_Path)

    If (File_Arrt And vb_Directory) = vb_Directory Then
        Get_Directory_Type = DirType_Folder
    Else
        Get_Directory_Type = DirType_File
    End If

    On Error GoTo 0: Exit Function
    '``````````````````````````````````````````````````'

'``````````````````````````````````````````````````````'
GS¦Absolute_Path:

    IsAbsolute_Path = False

    If Len(FileSystem_Path) > 3 Then
        If Left$(FileSystem_Path, 2) = "\\" Then
            IsAbsolute_Path = True: Return
        End If

        If Mid$(FileSystem_Path, 2, 1) = ":" Then
            If Mid(FileSystem_Path, 3, 1) = "\" Then
                IsAbsolute_Path = True: Return
            End If
        End If
    End If

    Return
'``````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````'
GT¦Directory_NotFound:

    On Error GoTo 0

    If Obj_FSO Is Nothing Then
        Set Obj_FSO = CreateObject("Scripting.FileSystemObject")
    End If

    With Obj_FSO
        If .FileExists(FileSystem_Path) Then
            Get_Directory_Type = DirType_File
        ElseIf .FolderExists(FileSystem_Path) Then
            Get_Directory_Type = DirType_Folder
        Else
            If Not Get_Directory_Type = DirType_Invalid Then
                Get_Directory_Type = DirType_NotFound
            End If
        End If
    End With
'```````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------'
End Function
'====================================================================='


'============================================================================='
Private Function Get_PathFileNameAndExtension( _
                 ByRef Full_Path As String _
        ) As Variant
'-----------------------------------------------------------------------------'

    '```````````````````````````'
    Dim File_Path      As String
    Dim File_Name      As String
    Dim File_Extension As String
    '```````````````````````````'
    Dim Full_Name      As String
    Dim Last_DotPos      As Long
    Dim Last_SlashPos    As Long
    '````````````````````````````'

    '````````````````````````````````````````````'
    Last_SlashPos = InStrRev(Full_Path, "\")
    File_Path = Left(Full_Path, Last_SlashPos - 1)
    '````````````````````````````````````````````'

    '````````````````````````````````````````````'
    Full_Name = Mid(Full_Path, Last_SlashPos + 1)
    '````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    Last_DotPos = InStrRev(Full_Name, ".")
    File_Name = Left(Full_Name, Last_DotPos - 1)
    File_Extension = Mid(Full_Name, Last_DotPos + 1)
    '```````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Get_PathFileNameAndExtension = Array(File_Path, File_Name, File_Extension)
    '`````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------'
End Function
'============================================================================='


'=============================================================='
Private Function Get_ArrayDimension( _
                 ByRef Source_Array As Variant _
        ) As Long
'--------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Dimension_Count As Long, tmp_Val As Long
    '```````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    On Error Resume Next
    If IsArray(Source_Array) Then
        Do While True
            tmp_Val = UBound(Source_Array, Dimension_Count + 1)
            If Err.Number <> 0 Then Exit Do
            Dimension_Count = Dimension_Count + 1
        Loop
    Else
        Get_ArrayDimension = -1
        Exit Function
    End If
    On Error GoTo 0
    '``````````````````````````````````````````````````````````'

    '```````````````````````````````````'
    Get_ArrayDimension = Dimension_Count
    '```````````````````````````````````'

'--------------------------------------------------------------'
End Function
'=============================================================='


'================================================================================'
Private Function Recoding_Text( _
                 ByRef Source_Text As String, _
                 Optional ByVal Encoding As MSOffice_Type_Encoding = Type_UTF_8 _
        ) As Byte()
'--------------------------------------------------------------------------------'

    '```````````````````````````````````````````````'
    Dim Byte_Result() As Byte, Ptr_Source As LongPtr
    Dim i  As Long, J As Long, Symb_Code  As Long
    Dim Char_Set    As String, Length     As Long
    '```````````````````````````````````````````````'
    Static Obj_Stream As Object
    '```````````````````````````````````````````'
    
    '````````````````````````````````````'
    If LenB(Source_Text) = 0 Then
        ReDim Byte_Result(0 To 0) As Byte
        Recoding_Text = Byte_Result
        Exit Function
    End If
    '````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````````'
    Select Case Encoding
        Case 0: Char_Set = "ASCII"
            Length = Len(Source_Text)
            ReDim Byte_Result(0 To Length - 1) As Byte
            
            For i = 1 To Length
                Byte_Result(i - 1) = Asc(Mid$(Source_Text, i, 1)) And &HFF
            Next i
            
            Recoding_Text = Byte_Result: Exit Function
            
        Case 2, 3:
            Length = LenB(Source_Text): Ptr_Source = StrPtr(Source_Text)
            ReDim Byte_Result(0 To Length - 1) As Byte
            
            CopyMemory Byte_Result(0), ByVal Ptr_Source, Length
            Recoding_Text = Byte_Result: Exit Function
            
        Case 1: Char_Set = "UTF-8"
        Case Else: Char_Set = "UTF-8"
    End Select
    '`````````````````````````````````````````````````````````````````````'
    
    '``````````````````````````````````````````````````````````````````````````'
    If Obj_Stream Is Nothing Then Set Obj_Stream = CreateObject("ADODB.Stream")
    '``````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    With Obj_Stream
        On Error Resume Next: .Close: On Error GoTo 0
        
        .Mode = 3: .Type = 2: .Open
        .Charset = Char_Set
        .WriteText Source_Text
        .Position = 0: .Type = 1
        
        If Encoding = 1 Then
            .Position = 3
        ElseIf Encoding > 1 Then
            .Position = 2
        End If
        
        Byte_Result = .Read
    End With
    '````````````````````````````````````````````````'
    
    '``````````````````````````'
    Recoding_Text = Byte_Result
    '``````````````````````````'
    
'--------------------------------------------------------------------------------'
End Function
'================================================================================'


'===================================================================='
Private Function IsModule_Exists( _
                 ByRef Module_Name As String _
        ) As Boolean
'--------------------------------------------------------------------'

    '```````````````````````'
    Dim Obj_VBComp As Object
    '```````````````````````'

    '```````````````````````````````````````````````````````````````'
    On Error Resume Next

    Set Obj_VBComp = ThisWorkbook.VBProject.VBComponents(Module_Name)
    IsModule_Exists = (Err.Number = 0): Set Obj_VBComp = Nothing

    On Error GoTo 0
    '```````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------'
End Function
'===================================================================='


'=================================================================================================='
Private Function IsProcedure_Exists( _
                 ByRef Module_Name As String, _
                 ByRef Procedure_Name As String _
        ) As Boolean
'--------------------------------------------------------------------------------------------------'
    
    '`````````````````````````````````````````````````'
    Dim Obj_VBProject As Object, Obj_VBComp  As Object
    Dim Full_Code   As Variant, Trimmed_Line As String
    '`````````````````````````````````````````````````'
    Dim Line_Count    As Long, i As Long
    Dim Inx_1         As Long, Inx_2         As Long
    Dim Inx_3         As Long, Inx_4         As Long
    Dim Length_1      As Long, Length_2      As Long
    Dim Length_3      As Long, Length_4      As Long
    Dim SubString_1   As String, SubString_2 As String
    Dim SubString_3   As String, SubString_4 As String
    Dim SubString_5   As String, SubString_6 As String
    '`````````````````````````````````````````````````'
    Static Obj_RegExp As Object
    '`````````````````````````````````````````````````'
    
    '````````````````````````'
    IsProcedure_Exists = True
    '````````````````````````'
    
    '``````````````````````````````````````````````````````````````'
    Length_1 = Len("Sub "):      Length_2 = Len("Public Sub ")
    Length_3 = Len("Function "): Length_4 = Len("Public Function ")
    '``````````````````````````````````````````````````````````````'
    
    '`````````````````````````````````````````````````'
    SubString_1 = "Sub " & Procedure_Name
    SubString_2 = "Public Sub " & Procedure_Name
    SubString_3 = "Function " & Procedure_Name
    SubString_4 = "Public Function " & Procedure_Name
    SubString_5 = " " & Procedure_Name & "("
    SubString_6 = " " & Procedure_Name & " ("
    '`````````````````````````````````````````````````'
    
    '``````````````````````````````````````````````````````````````````````````````````````````````'
    Set Obj_VBProject = ThisWorkbook.VBProject
    Set Obj_VBComp = Obj_VBProject.VBComponents(Module_Name)

    With Obj_VBComp.CodeModule
        Line_Count = .CountOfLines
        If Line_Count = 0 Then IsProcedure_Exists = False: Exit Function
        Full_Code = .Lines(1, Line_Count)
    End With

    Full_Code = Replace(Full_Code, " _" & vbCrLf, vbNullString)
    GoSub GS¦Remove_DuplicateSpaces: Full_Code = Split(Full_Code, vbCrLf)
    
    For i = LBound(Full_Code, 1) To UBound(Full_Code, 1)
        Trimmed_Line = Full_Code(i)
        
        If Trimmed_Line <> "" And Left(Trimmed_Line, 1) <> "'" Then
            Inx_1 = InStr(1, Trimmed_Line, SubString_1, vbTextCompare)
            Inx_2 = InStr(1, Trimmed_Line, SubString_2, vbTextCompare)
            Inx_3 = InStr(1, Trimmed_Line, SubString_3, vbTextCompare)
            Inx_4 = InStr(1, Trimmed_Line, SubString_4, vbTextCompare)
            
            Select Case True
                Case Inx_1 > 0
                    If InStr(Length_1, Trimmed_Line, SubString_5, vbTextCompare) Then Exit Function
                    If InStr(Length_1, Trimmed_Line, SubString_6, vbTextCompare) Then Exit Function
                    
                Case Inx_2 > 0
                    If InStr(Length_2, Trimmed_Line, SubString_5, vbTextCompare) Then Exit Function
                    If InStr(Length_2, Trimmed_Line, SubString_6, vbTextCompare) Then Exit Function
                
                Case Inx_3 > 0
                    If InStr(Length_3, Trimmed_Line, SubString_5, vbTextCompare) Then Exit Function
                    If InStr(Length_3, Trimmed_Line, SubString_6, vbTextCompare) Then Exit Function
                
                Case Inx_4 > 0
                    If InStr(Length_4, Trimmed_Line, SubString_5, vbTextCompare) Then Exit Function
                    If InStr(Length_4, Trimmed_Line, SubString_6, vbTextCompare) Then Exit Function

            End Select
        End If
    Next i

    IsProcedure_Exists = False: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````'
GS¦Remove_DuplicateSpaces:
    
    If Obj_RegExp Is Nothing Then
        Set Obj_RegExp = CreateObject("VBScript.RegExp")
        With Obj_RegExp
            .Global = True: .Pattern = " {2,}"
        End With
    End If
    
    Full_Code = Obj_RegExp.Replace(Trim$(Full_Code), " ")
    
    Return
'````````````````````````````````````````````````````````'
    
'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'==================================================================='
Private Function Find_ProcedureStart( _
                 ByRef Code_Module As Object, _
                 ByVal Procedure_Name As String _
        ) As Long
'-------------------------------------------------------------------'

    '```````````````````````````````````````````````````````'
    Dim Line_Count As String, Line_Text As String, i As Long
    '```````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````'
    Line_Count = Code_Module.CountOfLines

    For i = 1 To Line_Count
        Line_Text = Code_Module.Lines(i, 1)
        If InStr(Line_Text, " " & Procedure_Name) > 0 Then
            If InStr(Line_Text, " " & Procedure_Name & " ") > 0 Then
                Find_ProcedureStart = i: Exit Function
            End If
            
            If InStr(Line_Text, " " & Procedure_Name & "(") > 0 Then
                Find_ProcedureStart = i: Exit Function
            End If
        End If
        
        If InStr(Line_Text, Procedure_Name) > 0 Then
            If InStr(Line_Text, Procedure_Name & " ") = 1 Then
                Find_ProcedureStart = i: Exit Function
            End If
            
            If InStr(Line_Text, Procedure_Name & "(") = 1 Then
                Find_ProcedureStart = i: Exit Function
            End If
        End If
    Next i
    '```````````````````````````````````````````````````````````````'
    
'-------------------------------------------------------------------'
End Function
'==================================================================='


'==================================================================='
Private Function Find_ProcedureEnd( _
                 ByRef Code_Module As Object, _
                 ByVal Start_Line As Long _
        ) As Long
'-------------------------------------------------------------------'

    '`````````````````````````````````'
    Dim Line_Text As String, i As Long
    '`````````````````````````````````'
    
    '```````````````````````````````````````````````````````````````'
    For i = Start_Line + 1 To Code_Module.CountOfLines
        Line_Text = Code_Module.Lines(i, 1)
        If UCase$(Left$(Trim$(Line_Text), 7)) = "END SUB" Or _
           UCase$(Left$(Trim$(Line_Text), 12)) = "END FUNCTION" Then
            Find_ProcedureEnd = i: Exit Function
        End If
    Next i
    '```````````````````````````````````````````````````````````````'
    
'-------------------------------------------------------------------'
End Function
'==================================================================='


'========================================================================'
Private Function Find_EncryptedBlock( _
                 ByRef Code_Module As Object, _
                 ByRef Procedure_Name As String, _
                 ByRef GUID_ID As String, _
                 ByRef Start_Line As Long, _
                 ByRef End_Line As Long, _
                 ByRef Encrypted_Lines As Collection _
        ) As Boolean
'------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````'
    Dim Line_Count As Long, Line_Text As String, i As Long
    '`````````````````````````````````````````````````````'
    
    '````````````````````````````````````````````````````````````````````'
    Line_Count = Code_Module.CountOfLines
    
    For i = 1 To Line_Count
        Line_Text = Code_Module.Lines(i, 1)
        If InStr(Line_Text, "// Procedure: " & _
                                Procedure_Name & " " & GUID_ID) > 0 Then
            Start_Line = i
            Do
                i = i + 1: Line_Text = Code_Module.Lines(i, 1)
                Encrypted_Lines.Add Line_Text
                If InStr(Line_Text, "// End " & GUID_ID) > 0 Then
                    End_Line = i: Find_EncryptedBlock = True
                    Exit Function
                End If
            Loop While i <= Line_Count
        End If
    Next i
    '````````````````````````````````````````````````````````````````````'
    
'------------------------------------------------------------------------'
End Function
'========================================================================'


'==========================================================================================='
Private Function Show_ErrorMessage_Immediate( _
                 ByRef Error_Message As String, _
                 Optional ByRef Header_Message As Variant = vbNullChar, _
                 Optional ByRef WinAPI_DllName As String = vbNullString, _
                 Optional ByVal Show_Immediate As Boolean = False _
        ) As String
'-------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````'
    Dim RT_Error_Number As Long, RT_Error_Description    As String
    Dim Len_Description As Long, Len_Message&, Len_Borders As Long
    Dim WB_Name As String, Debug_Borders As String
    '`````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    RT_Error_Number = Err.Number: RT_Error_Description = Err.Description
    RT_Error_Description = Replace(RT_Error_Description, ChrW$(10), vbNullChar)
    '`````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    If Show_Immediate Then
        With Application.VBE.Windows("Immediate")
            .Visible = True: .SetFocus
        End With
    End If
    '````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Len_Description = Len(RT_Error_Description): Len_Message = Len(Error_Message)
    Len_Borders = IIf(Len_Description > Len_Message, Len_Description, Len_Message)

    Debug_Borders = Debug_Borders & "+"
    Debug_Borders = Debug_Borders & VBA.String$(18 + Len_Borders, "-")
    Debug_Borders = Debug_Borders & "+"
    '`````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    If Not Header_Message = vbNullChar Then
        If TypeName(Header_Message) = "String" Then
            WB_Name = Header_Message
        Else
            WB_Name = Header_Message.Name: If Len(WB_Name) = 0 Then WB_Name = "Unknown Book"
        End If
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Debug.Print "..." & vbNewLine

    If CBool(RT_Error_Number) Then
        Debug.Print "RunTime Error (" & RT_Error_Number & ") > " & WB_Name
        Debug.Print Debug_Borders & vbNewLine & "| Description:   | " & RT_Error_Description
    Else
        Debug.Print "Custom Exception" & " > " & WB_Name & vbNewLine & Debug_Borders
    End If

    Debug.Print "| Error_Message: | " & Error_Message & vbNewLine & Debug_Borders
    '```````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------'
End Function
'==========================================================================================='


'=================================================================================='
Private Function Get_ErrorMessage_WinAPI( _
                 ByVal Error_ID As Long _
        ) As String
'----------------------------------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim WinAPI_Result As Long, Buffer As String
    '``````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    Buffer = String(256, vbNullChar)

    WinAPI_Result = FormatMessage( _
                    FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                    0, Error_ID, LANG_DEFAULT, Buffer, Len(Buffer), 0)
    '``````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If Not WinAPI_Result = 0 Then
        Get_ErrorMessage_WinAPI = Left$(Buffer, InStr(Buffer, vbNewLine) - 1)
    Else
        Get_ErrorMessage_WinAPI = "Не удалось получить сообщение об ошибке"
    End If
    '`````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------'
End Function
'=================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================================='
Private Function MacroSecurity_MacroParameters_Assign( _
                 ByVal Type_Document As MSOffice_Type_Document_Group, _
                 Optional ByVal Macro_Privileges As Security_ManagementCenter_Privileges = Without_PrivilegesChanges, _
                 Optional ByVal Macro_AccessObjectModel As Security_ManagementCenter_AccessObjectModel = Without_AOMChanges, _
                 Optional ByVal Macro_ForcedChange As Boolean = True _
        ) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-----------------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````'
    Dim MSOffice_App     As Object
    Dim Registry_Section As String, Section_Data As Registry_SectionData
    Dim Handle As LongPtr, API_lpData As Long, Native_Process As Boolean
    '```````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    MacroSecurity_MacroParameters_Assign = False
    '```````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````'
    If Macro_Privileges = Without_PrivilegesChanges Then
        If Macro_AccessObjectModel = Without_AOMChanges Then Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````'

    '`````````````````````````````'
    Call Init_VBD_Kit_Cryptography
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document

        Case MSOffice_Type_Document_Group.mso_Access
             Registry_Section = Registry_Section & REGISTRY_SECTION_ACCESS_SECURITY

        Case MSOffice_Type_Document_Group.mso_Excel
             Registry_Section = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY

        Case MSOffice_Type_Document_Group.mso_Outlook
             Registry_Section = Registry_Section & REGISTRY_SECTION_OUTLOOK_SECURITY

        Case MSOffice_Type_Document_Group.mso_PowerPoint
             Registry_Section = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY

        Case MSOffice_Type_Document_Group.mso_Word
             Registry_Section = Registry_Section & REGISTRY_SECTION_WORD_SECURITY

    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_ALL_ACCESS)
    Handle = Section_Data.Handle: If Handle = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_Document

        Case mso_Access
             If Macro_Privileges = Without_PrivilegesChanges Then Exit Function
             GoSub GS¦Change_Registry_VBAWarnings

        Case mso_Excel
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then Native_Process = True

        Case mso_PowerPoint
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_PowerPoint Then Native_Process = True

        Case mso_Word
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_Word Then Native_Process = True

        Case mso_Outlook
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_Level
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_DontTrustInstalledFiles

    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    If Macro_ForcedChange Then
        If MacroSecurity_MacroParameters_Assign Then
            If Native_Process Then GoSub GS¦Update_VBE_Security
        End If
    End If

    Call RegCloseKey(Handle): Exit Function
    '``````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````'
GS¦Update_VBE_Security:

    Set MSOffice_App = Application

    Call SendKeys("{ENTER}")
    Call MSOffice_App.CommandBars.ExecuteMso("MacroSecurity")

    Return
'``````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_VBAWarnings:

    Call RegQueryValueEx(Handle, "VBAWarnings", 0&, REG_DWORD, API_lpData, 4&)

    If API_lpData <> Macro_Privileges Then
        MacroSecurity_MacroParameters_Assign = RegSetValueEx(Handle, "VBAWarnings", 0&, _
                                             REG_DWORD, Macro_Privileges, 4&) = 0&
    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_AccessVBOM:

    Call RegQueryValueEx(Handle, "AccessVBOM", 0&, REG_DWORD, API_lpData, 4&)

    If API_lpData <> Macro_AccessObjectModel Then
        MacroSecurity_MacroParameters_Assign = RegSetValueEx(Handle, "AccessVBOM", 0&, _
                                             REG_DWORD, Macro_AccessObjectModel, 4&) = 0&
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_DontTrustInstalledFiles:

    Call RegQueryValueEx(Handle, "DontTrustInstalledFiles", 0&, REG_DWORD, API_lpData, 4&)

    If API_lpData <> Macro_AccessObjectModel Then
        MacroSecurity_MacroParameters_Assign = RegSetValueEx(Handle, "DontTrustInstalledFiles", 0&, _
                                              REG_DWORD, Macro_AccessObjectModel, 4&) = 0&
    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_Level:

    Call RegQueryValueEx(Handle, "Level", 0&, REG_DWORD, API_lpData, 4&)

    If API_lpData <> Macro_Privileges Then
        MacroSecurity_MacroParameters_Assign = RegSetValueEx(Handle, "Level", 0&, _
                                              REG_DWORD, Macro_Privileges, 4&) = 0&
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================================='


'======================================================================================================'
Private Function Get_Section_Data( _
                 ByVal Registry_HKey As Registry_HKeys, _
                 ByRef Registry_Section As String, _
                 ByVal Access_Rights As Registry_AccessTypes _
        ) As Registry_SectionData
'------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, WinAPI_Result As Long, Error_Message As String
    Dim lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcValues As Long
    Dim lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long
    Dim lpftLastWriteTime As WinAPI_FileTime, lpftSystemTime As WinAPI_SystemTime
    '````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````'
    WinAPI_Result = RegOpenKeyEx(Registry_HKey, Registry_Section, 0&, Access_Rights, Handle)
    '``````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case WinAPI_Result
        Case ERROR_SUCCESS
        Case ERROR_FILE_NOT_FOUND

            Select Case True
                Case Registry_Section Like "*\Security\ProtectedView": GoTo GS¦Fill_Data
            End Select

            Error_Message = Get_ErrorMessage_WinAPI(ERROR_FILE_NOT_FOUND) & _
                            " (Проверьте аргумент ""Registry_Section"")"

            Show_ErrorMessage_Immediate Error_Message, _
                                        "Ошибка при вызове WinAPI", "advapi32.dll_RegOpenKeyEx"
            Exit Function

        Case ERROR_ACCESS_DENIED
            Error_Message = Get_ErrorMessage_WinAPI(ERROR_ACCESS_DENIED) & _
                            " (В данный момент невозможно получить требуемый доступ к разделу реестру)"

            Show_ErrorMessage_Immediate Error_Message, _
                                        "Ошибка при вызове WinAPI", "advapi32.dll_RegOpenKeyEx"
            Exit Function

        Case Else
            Error_Message = Get_ErrorMessage_WinAPI(WinAPI_Result)

            Show_ErrorMessage_Immediate Error_Message, _
                                        "Ошибка при вызове WinAPI", "advapi32.dll_RegOpenKeyEx"
            Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    WinAPI_Result = RegQueryInfoKey(Handle, vbNullString, 0&, 0&, lpcSubKeys, lpcbMaxSubKeyLen, _
                                    ByVal 0&, lpcValues, lpcbMaxValueNameLen, lpcbMaxValueLen, _
                                    ByVal 0&, lpftLastWriteTime)

    If Not WinAPI_Result = ERROR_SUCCESS Then
        Error_Message = Get_ErrorMessage_WinAPI(WinAPI_Result)

        Show_ErrorMessage_Immediate Error_Message, _
                                    "Ошибка при вызове WinAPI", "advapi32.dll_RegQueryInfoKey"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Fill_Data:

    Get_Section_Data.Handle = Handle
    Get_Section_Data.Section_HKey = Registry_HKey
    Get_Section_Data.Section_Path = Registry_Section

    Get_Section_Data.Keys_Count = lpcValues
    Get_Section_Data.Keys_MaxLen_Name = lpcbMaxValueNameLen
    Get_Section_Data.Keys_MaxLen_Value = lpcbMaxValueLen

    Get_Section_Data.Section_AccessRight = Access_Rights
    Get_Section_Data.Section_Count = lpcSubKeys
    Get_Section_Data.Section_MaxLen_Name = lpcbMaxSubKeyLen

    WinAPI_Result = FileTimeToSystemTime(lpftLastWriteTime, Get_Section_Data.Section_LastWriteTime)
'``````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'======================================================================================================'
Private Function Show_InfoMessage_UnixNotSupported()
'------------------------------------------------------------------------------------------------------'
    
    '```````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox "Данный программный модуль не поддерживает работу в UNIX-подобных системах (MacOS)! " & _
       String$(2, vbNewLine) & "Реализована поддержка только в Windows_NT системах!" & vbNewLine & _
      "Дальнейшая работа функции будет прервана!", vbInformation, "[Cyber_Automation]"
    '```````````````````````````````````````````````````````````````````````````````````````````````'
    
'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'--------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------'

#End If

'========================================================================================================'
Private Sub Setup_Configuration()
'--------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````'
    Dim MSOffice_App As Object, Flag_Update As Boolean
    '`````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next

    If Len(ThisWorkbook.VBProject.Name) = 0 Then
        MsgBox "Для автоматического определения конфигурации сборки требуется " & _
               "поставить галочку в настройках макросов: " & vbNewLine & _
               """Доверять доступ к объектной модели проектов VBA""", vbInformation, "[Cyber_Automation]"

        #If Windows_NT Then
            GoSub GS¦Open_VBE_Security
            If Len(ThisWorkbook.VBProject.Name) = 0 Then On Error GoTo 0: Exit Sub
        #Else
            On Error GoTo 0: Exit Sub
        #End If
    End If

    On Error GoTo 0
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    Flag_Update = False: GoTo GT¦Change_ConstantByGUID: Exit Sub
    '```````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````'
GS¦Open_VBE_Security:

    Set MSOffice_App = Application
    Call MSOffice_App.CommandBars.ExecuteMso("MacroSecurity")

    Return
'````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````'
GS¦VBInfo_SuccessfulCodeExecution:

    MsgBox "Конфигурация сборки компонента успешно изменена!", _
            vbInformation, "[Cyber_Automation]"

    Return
'````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GT¦Change_ConstantByGUID:

    #If Windows_NT Then
        If Update_ConstantByGUID(GUID_VBComponent, "BuildMode_WindowsNT_Only", True) Then
           Update_ConstantByGUID GUID_VBComponent, "BuildMode_MacOS_Only", False
           Flag_Update = True
        End If
    #Else
        If Update_ConstantByGUID(GUID_VBComponent, "BuildMode_WindowsNT_Only", False) Then
           Update_ConstantByGUID GUID_VBComponent, "BuildMode_MacOS_Only", True
           Flag_Update = True
        End If
    #End If

    #If x64_Soft Then
        If Update_ConstantByGUID(GUID_VBComponent, "BuildMode_x64Only", True) Then
           Update_ConstantByGUID GUID_VBComponent, "BuildMode_x32Only", False
           If Flag_Update Then GoSub GS¦VBInfo_SuccessfulCodeExecution
        End If
    #Else
        If Update_ConstantByGUID(GUID_VBComponent, "BuildMode_x64Only", False) Then
           Update_ConstantByGUID GUID_VBComponent, "BuildMode_x32Only", True
           If Flag_Update Then GoSub GS¦VBInfo_SuccessfulCodeExecution
        End If
    #End If
'`````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Sub
'========================================================================================================'


'================================================================================'
Private Function Update_ConstantByGUID( _
                 ByRef Target_GUID As String, _
                 ByRef Const_Name As String, _
                 ByVal Const_Value As Variant _
        ) As Boolean
'--------------------------------------------------------------------------------'

    '`````````````````````````````'
    Dim Obj_VBComp       As Object
    Dim Obj_CodeModule   As Object
    Dim i As Long, Line  As String
    Dim Inx_Pos As Long
    Dim New_Line         As String
    Dim Const_NameInLine As String
    '`````````````````````````````'

    '````````````````````````````'
    Update_ConstantByGUID = False
    '````````````````````````````'

    '``````````````````````````````````````````````'
    Set Obj_VBComp = Find_ModuleByGUID(Target_GUID)
    If Obj_VBComp Is Nothing Then Exit Function
    '``````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    Set Obj_CodeModule = Obj_VBComp.CodeModule

    For i = 1 To Obj_CodeModule.CountOfLines
        Line = Obj_CodeModule.Lines(i, 1)
        If Left$(Line, 7) = "#Const " Then
            Inx_Pos = InStr(Line, "=")
            If Inx_Pos > 0 Then
                Const_NameInLine = Trim(Mid$(Line, 8, Inx_Pos - 8))
                If Const_NameInLine = Const_Name Then
                    New_Line = "#Const " & Const_Name & " = " & CStr(Const_Value)
                    Obj_CodeModule.ReplaceLine i, New_Line
                    Update_ConstantByGUID = True
                    Exit Function
                End If
            End If
        End If
    Next i
    '````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    MsgBox "Обнаружено изменение структуры конфигурации сборки компонента! " _
           & vbNewLine & _
           "Попытка автоматического изменения конфигурации прервана!", _
           vbCritical, "[Cyber_Automation]"
    '``````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------'
End Function
'================================================================================'


'================================================================================'
Private Function Find_ModuleByGUID( _
                 ByRef Target_GUID As String _
        ) As Object
'--------------------------------------------------------------------------------'

    '````````````````````````````'
    Dim Obj_VBProject   As Object
    Dim Obj_VBComp      As Object
    Dim Obj_CodeModule  As Object
    Dim i As Long, Line As String
    Dim Inx_Pos As Long
    Dim GUID_Value As String
    '````````````````````````````'

    '`````````````````````````````````````````'
    Set Obj_VBProject = ThisWorkbook.VBProject
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    For Each Obj_VBComp In Obj_VBProject.VBComponents
        Set Obj_CodeModule = Obj_VBComp.CodeModule
        For i = 1 To Obj_CodeModule.CountOfLines
            Line = Obj_CodeModule.Lines(i, 1)
            If InStr(Line, "GUID_VBComponent") > 0 And InStr(Line, "=") > 0 Then
                Inx_Pos = InStr(Line, "=")
                If InStr(Line, "'") > Inx_Pos Or InStr(Line, "'") = 0 Then
                    GUID_Value = Trim$(Mid$(Line, Inx_Pos + 1))
    
                    If Left$(GUID_Value, 1) = """" Then
                        GUID_Value = Mid$(GUID_Value, 2)
                    End If
    
                    If Right$(GUID_Value, 1) = """" Then
                        GUID_Value = Left$(GUID_Value, Len(GUID_Value) - 1)
                    End If
    
                    If GUID_Value = Target_GUID Then
                        Set Find_ModuleByGUID = Obj_VBComp: Exit Function
                    End If
                End If
            End If
        Next i
    Next Obj_VBComp
    '````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------'
End Function
'================================================================================'

'--------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------'
