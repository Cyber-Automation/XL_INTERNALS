Attribute VB_Name = "VBD_Kit_Hashing"
' | ========================================================================================================================= | '
' | __________  __________________________     ____________  _____________________  ___________________________________   __  | '
' | __  ____/ \/ /__  __ )__  ____/__  __ \    ___    |_  / / /__  __/_  __ \__   |/  /__    |__  __/___  _/_  __ \__  | / /  | '
' | _  /    __  /__  __  |_  __/  __  /_/ /    __  /| |  / / /__  /  _  / / /_  /|_/ /__  /| |_  /   __  / _  / / /_   |/ /   | '
' | / /___  _  / _  /_/ /_  /___  _  _, _/     _  ___ / /_/ / _  /   / /_/ /_  /  / / _  ___ |  /   __/ /  / /_/ /_  /|  /    | '
' | \____/  /_/  /_____/ /_____/  /_/ |_|______/_/  |_\____/  /_/    \____/ /_/  /_/  /_/  |_/_/    /___/  \____/ /_/ |_/     | '
' |                                     _/_____/                                                                              | '
' | ========================================================================================================================= | '

' +-[MODULE: VBD_Kit_Hashing]----------------------------------------------------------------+
' |                                                                                          |
' | [ENGINEER]: Zeus_0x01                                                                    |
' | [TELEGRAM]: @Zeus_0x01 (Public Name)                                                     |
' | [DESCRIPTION]: Реализация множества алгоритмов для вычисления хеш-суммы различных данных |
' |                                                                                          |
' +------------------------------------------------------------------------------------------+

' // <copyright file="VBD_Kit_Hashing.bas" division="Cyber_Automation">
' // (C) Copyright 2023 Zeus_0x01 "{BDED7A00-EC23-40CD-AA90-7C952ECF7621}"
' // </copyright>

'-----------------------------------------------------------------------'
' // Implemented Functionality (Реализованный функционал):
'    Forum -  https://www.script-coding.ru/threads/vbd_kit_hashing.199/
'    GitHub - https://github.com/Cyber-Automation/XL_INTERNALS/
'-----------------------------------------------------------------------'
' // Release_Version (Версия компонента) - [01.01]
'-----------------------------------------------------------------------'

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
Private Const GUID_VBComponent As String = "{BDED7A00-EC23-40CD-AA90-7C952ECF7621}"
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

'-------------------------------------------'
Public Enum VBProject_Crypt_HashingProvider
    Microsoft_NET = &H1
    Microsoft_WinAPI = &H2
    Native_VBACode = &H3
End Enum
'-------------------------------------------'

'-------------------------------------------'
Public Enum VBProject_Crypt_HashingAlgorithm
    MD_5 = &H1
    SHA_1 = &H2
    SHA_256 = &H3
    SHA_384 = &H4
    SHA_512 = &H5
End Enum
'-------------------------------------------'

'-------------------------------------------'
Public Enum MSOffice_Type_ContentFormat_HS
    MSDI_Undefined_HS = &HFFFFFFFC
    MSDI_Null_HS = &HFFFFFFFD
    MSDI_NullString_HS = &HFFFFFFFE
    MSDI_Nothing_HS = &HFFFFFFFF
    MSDI_Empty_HS = &H0
    MSDI_Bool_HS = &H1
    MSDI_Int8_HS = &H2
    MSDI_Int16_HS = &H4
    MSDI_Int32_HS = &H8
    MSDI_Int64_HS = &H10
    MSDI_FloatPoint32_HS = &H20
    MSDI_FloatPoint64Cur_HS = &H40
    MSDI_FloatPoint64Dbl_HS = &H56
    MSDI_FloatPoint112_HS = &H80
    MSDI_Date_HS = &H100
    MSDI_String_HS = &H200
    MSDI_Range_HS = &H400
    MSDI_Array_HS = &H800
    MSDI_Object_HS = &H1000
    MSDI_Collection_HS = &H2000
    MSDI_Dictionary_HS = &H4000
    MSDI_File_HS = &H8000
    MSDI_Folder_HS = &H10000
    MSDI_VBComponent_HS = &H20000
    MSDI_VBProject_HS = &H40000
    MSDI_UserDefType_HS = &H80000
    MSDI_NonExistent_Directory_HS = &H100000
    
    ' {
        MSDI_Procedure_HS = &H120000
        MSDI_AutoDetect_HS = &HFFFFFFAD
    ' }
End Enum
'-------------------------------------------'

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
Private Enum MSOffice_Type_Encoding
    Type_ASCII = &H0
    Type_UTF_8 = &H1
    Type_UTF_16 = &H2
    Type_Unicode = &H3
End Enum
'-----------------------------------'

'---------------------------------------'
Private Enum MSOffice_Type_Building
    Type_MSOffice_Build_Undefined = &H0
    Type_MSOffice_Build_NotSupport = &H1
    Type_MSOffice_2016_32Bit = &H2
    Type_MSOffice_2016_64Bit = &H3
    Type_MSOffice_2019_365_32Bit = &H4
    Type_MSOffice_2019_365_64Bit = &H5
End Enum
'---------------------------------------'

'---------------------------------------'
Private Enum MSOffice_Type_Application
    Type_MSOffice_App_Undefined = &H0
    Type_MSOffice_App_NotSupport = &H1
    Type_MSOffice_Excel = &H2
    Type_MSOffice_PowerPoint = &H3
    Type_MSOffice_Word = &H4
    Type_MSOffice_Access = &H5
    Type_MSOffice_Outlook = &H5
End Enum
'---------------------------------------'

'----------------------------------------'
Private Enum MSOffice_Type_Document_Group
    mso_Excel = &H2
    mso_PowerPoint = &H3
    mso_Word = &H4
    mso_Access = &H5
    mso_Outlook = &H6
End Enum
'----------------------------------------'

'--------------------------------------'
Private Enum FileSystem_Directory_Type
    DirType_File = &H0
    DirType_Folder = &H1
    DirType_Invalid = &HFFFFFFFF
    DirType_NotFound = &HFFFFFFFE
End Enum
'--------------------------------------'

'--------------------------------------'
Private Enum FileSystem_Format_SizeFile
    SizeFormat_Byte = &H0
    SizeFormat_KByte = &H1
    SizeFormat_MByte = &H2
    SizeFormat_GByte = &H3
End Enum
'--------------------------------------'

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

'--------------------------------'
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
'--------------------------------'

'--------------------------------'
Private Enum Registry_ValueTypes
    REG_SZ = &H1
    REG_BINARY = &H3
    REG_DWORD = &H4
    REG_QWORD = &HB
    REG_MULTI_SZ = &H7
    REG_EXPAND_SZ = &H2
End Enum
'--------------------------------'

'--------------------------------'
Private Type WinAPI_LargeInteger
    Low_Part  As Long
    High_Part As Long
End Type
'--------------------------------'

'----------------------------'
Private Type WinAPI_FileTime
    tm_LowDateTime  As Long
    tm_HighDateTime As Long
End Type
'----------------------------'

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

'-----------------------------------------------------------------------------------'
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
            Function GetFileSizeEx Lib "kernel32.dll" ( _
                     ByVal hFile As LongPtr, _
                     ByRef lpFileSize As WinAPI_LargeInteger _
            ) As Long

    Private Declare PtrSafe _
            Function CloseHandle Lib "kernel32.dll" ( _
                     ByVal hObject As LongPtr _
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

#Else

    ' // Old MS Office or MacOS

#End If
'-----------------------------------------------------------------------------------'

'--------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (VBE7.dll)

    Private Declare PtrSafe _
            Function ArrPtr Lib "VBE7.dll" Alias "VarPtr" ( _
                     ByRef Ptr() As Any _
            ) As LongPtr

#Else

    ' // Old MS Office or MacOS

#End If
'--------------------------------------------------------------'

'----------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Bcrypt.dll)

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
                     ByVal dwFlags As Long _
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

#Else

    ' // Old MS Office or MacOS

#End If
'----------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Advapi32.dll)

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
'------------------------------------------------------------------------------------'

'---------------------------------------------------------'
Private Const INVALID_HANDLE_VALUE As LongPtr = &HFFFFFFFF
'---------------------------------------------------------'

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

'------------------------------------------------------------------'
Private Glb_MSOffice_Type_Building    As MSOffice_Type_Building
Private Glb_MSOffice_Type_Application As MSOffice_Type_Application
'------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------'
'------------------------------------------------------------------'
Private MD5_lngTrack             As Long
Private MD5_ArrLongConversion(4) As Long
Private MD5_ArrSplit64(63)       As Byte

Private Const MD5_OFFSET_4 = 4294967296#, MD5_MAXINT_4 = 2147483647
Private Const MD5_S11 = 7, MD5_S12 = 12, MD5_S13 = 17, MD5_S14 = 22
Private Const MD5_S24 = 20, MD5_S31 = 4, MD5_S32 = 11, MD5_S33 = 16
Private Const MD5_S21 = 5, MD5_S22 = 9, MD5_S23 = 14, MD5_S34 = 23
Private Const MD5_S41 = 6, MD5_S42 = 10, MD5_S43 = 15, MD5_S44 = 21
'------------------------------------------------------------------'

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

'--------------------------------------------------------------------------------------------'
Private SHA256_m_lOnBits(30) As Long
Private SHA256_m_l2Power(30) As Long
Private SHA256_K(63)         As Long

Private Const SHA256_BITS_TO_A_BYTE  As Long = 8&
Private Const SHA256_BYTES_TO_A_WORD As Long = 4&
Private Const SHA256_BITS_TO_A_WORD  As Long = SHA256_BYTES_TO_A_WORD * SHA256_BITS_TO_A_BYTE
'--------------------------------------------------------------------------------------------'

'-------------------------------------------------------------'
Private Const SHA512_LNG_BLOCKSZ As Long = 128&
Private Const SHA512_LNG_ROUNDS  As Long = 80&
Private Const SHA512_LNG_POW2_1  As Long = 2& ^ 1&
Private Const SHA512_LNG_POW2_2  As Long = 2& ^ 2&
Private Const SHA512_LNG_POW2_3  As Long = 2& ^ 3&
Private Const SHA512_LNG_POW2_4  As Long = 2& ^ 4&
Private Const SHA512_LNG_POW2_5  As Long = 2& ^ 5&
Private Const SHA512_LNG_POW2_6  As Long = 2& ^ 6&
Private Const SHA512_LNG_POW2_7  As Long = 2& ^ 7&
Private Const SHA512_LNG_POW2_8  As Long = 2& ^ 8&
Private Const SHA512_LNG_POW2_9  As Long = 2& ^ 9&
Private Const SHA512_LNG_POW2_12 As Long = 2& ^ 12&
Private Const SHA512_LNG_POW2_13 As Long = 2& ^ 13&
Private Const SHA512_LNG_POW2_14 As Long = 2& ^ 14&
Private Const SHA512_LNG_POW2_17 As Long = 2& ^ 17&
Private Const SHA512_LNG_POW2_18 As Long = 2& ^ 18&
Private Const SHA512_LNG_POW2_19 As Long = 2& ^ 19&
Private Const SHA512_LNG_POW2_22 As Long = 2& ^ 22&
Private Const SHA512_LNG_POW2_23 As Long = 2& ^ 23&
Private Const SHA512_LNG_POW2_24 As Long = 2& ^ 24&
Private Const SHA512_LNG_POW2_25 As Long = 2& ^ 25&
Private Const SHA512_LNG_POW2_26 As Long = 2& ^ 26&
Private Const SHA512_LNG_POW2_27 As Long = 2& ^ 27&
Private Const SHA512_LNG_POW2_28 As Long = 2& ^ 28&
Private Const SHA512_LNG_POW2_29 As Long = 2& ^ 29&
Private Const SHA512_LNG_POW2_30 As Long = 2& ^ 30&
Private Const SHA512_LNG_POW2_31 As Long = &H80000000

Private SHA512_LNG_K(0& To 2& * SHA512_LNG_ROUNDS - 1&) As Long
Private SHA512_m_bNoIntegerOverflowChecks            As Boolean

Private Type SHA512_SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As LongPtr
    cElements  As Long
    lLbound    As Long
End Type

Private Type SHA512_ArrayLong16
    Item(0 To 15) As Long
End Type

Private Type SHA512_ArrayLong32
    Item(0 To 31) As Long
End Type

Public Type SHA512_CryptoSha512Context
    State      As SHA512_ArrayLong16
    Block      As SHA512_ArrayLong32
    bytes()    As Byte
    ArrayBytes As SHA512_SAFEARRAY1D
    NPartial   As Long
    NInput     As Currency
    BitSize    As Long
End Type
'-------------------------------------------------------------'
'--------------------------------------------------------------------------------------------'


'=========================================================================================================='
Public Function Get_HashSumm_Data( _
                ByRef Source_Data As Variant, _
                Optional ByVal Explicit_DataType As MSOffice_Type_ContentFormat_HS = MSDI_AutoDetect_HS, _
                Optional ByVal Hashing_Algorithm As VBProject_Crypt_HashingAlgorithm = SHA_256, _
                Optional ByVal Hashing_Method As VBProject_Crypt_HashingProvider = Microsoft_WinAPI, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal dw_Reserved_2 As Long = 0& _
        ) As Variant
'----------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-get_hashsumm_data-vbd_kit_hashing-bas.292/
'``````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - Hash_OnlyCode      | Хешировать только исходный код (без учёта комментариев и отступов)
' dw_Reserved_2 - Include_SubFolders | Хешировать файлы в папках с анализом подпапок
'----------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````'
    Dim Data_Bytes()  As Byte, Data_Type      As MSOffice_Type_ContentFormat_HS
    Dim File_Path     As String, HS_Algorithm As String, Error_Message As String
    Dim Obj_VBProject As Object, Obj_VBComponent As Object, VBProject  As Object
    Dim Vector_HashSumm() As String, Count_VBComponent As Long, Inx    As Long
    Dim Flag_SuccessfulHashSummCalculation  As Boolean
    '```````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Get_HashSumm_Data = vbNullString
    '```````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Hashing
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then
        Exit Function
    End If

    If Glb_MSOffice_Type_Application = Type_MSOffice_App_NotSupport Then
        Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    If Explicit_DataType = MSDI_AutoDetect_HS Then
        Data_Type = Get_SemanticDataType(Source_Data)
    Else
        Data_Type = Explicit_DataType
    End If
    '````````````````````````````````````````````````'

    '```````````````````````````````````'
    Select Case Hashing_Algorithm
        Case 0: HS_Algorithm = "MD5"
        Case 1: HS_Algorithm = "SHA1"
        Case 2: HS_Algorithm = "SHA256"
        Case 3: HS_Algorithm = "SHA384"
        Case 4: HS_Algorithm = "SHA512"
    End Select
    '```````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    Select Case Data_Type
        Case MSDI_NonExistent_Directory_HS: GoSub GS¦Exception_NonExistent_Directory
        Case MSDI_NullString_HS:            GoSub GS¦Exception_NullString
        Case MSDI_Range_HS, MSDI_Array_HS:  GoSub GS¦Hashing_Array
        Case MSDI_Bool_HS To MSDI_Date_HS:  GoSub GS¦Hashing_Number
        Case MSDI_String_HS:                GoSub GS¦Hashing_String
        Case MSDI_File_HS:                  GoSub GS¦Hashing_File
        Case MSDI_Folder_HS:                GoSub GS¦Hashing_Folder
        Case MSDI_VBComponent_HS:           GoSub GS¦Hashing_VBComponent
        Case MSDI_VBProject_HS:         GoSub GS¦Hashing_VBProject
        
        Case Else
            Error_Message = "Тип данных не поддерживается: " & TypeName(Source_Data)
            Call Show_ErrorMessage_Immediate(Error_Message, "Несоответствие типов")
            Get_HashSumm_Data = vbNullString

    End Select

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````'
GS¦Hashing_Data:

    Select Case Hashing_Method
        Case Microsoft_WinAPI:
            Get_HashSumm_Data = Hashing_WinAPI(Data_Bytes, HS_Algorithm)
            
        Case Microsoft_NET:
            If Check_NetFramework_IsLoaded Then
                Get_HashSumm_Data = Hashing_NET(Data_Bytes, HS_Algorithm)
            End If
        
        Case Hashing_Method:
            Get_HashSumm_Data = Hashing_NativeCode(Data_Bytes, HS_Algorithm)
        
    End Select

    Return
'```````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GS¦Hashing_Number:

    Data_Bytes = Recoding_Text(CStr(Source_Data)): GoSub GS¦Hashing_Data
    
    Return
'````````````````````````````````````````````````````````````````````````'
    
'``````````````````````````````````````````````````````````````````````````````'
GS¦Hashing_String:

    If LenB(Source_Data) = 0 Then
        Select Case Hashing_Algorithm
            Case 0: Get_HashSumm_Data = "d41d8cd98f00b204e9800998ecf8427e"
            Case 1: Get_HashSumm_Data = "da39a3ee5e6b4b0d3255bfef95601890" & _
                                        "afd80709"
            Case 2: Get_HashSumm_Data = "e3b0c44298fc1c149afbf4c8996fb924" & _
                                        "27ae41e4649b934ca495991b7852b855"
            Case 3: Get_HashSumm_Data = "38b060a751ac96384cd9327eb1b1e36a" & _
                                        "21fdb71114be07434c0cc7bf63f6e1da" & _
                                        "274edebfe76f65fbd51ad2f14898b95b"
            Case 4: Get_HashSumm_Data = "cf83e1357eefb8bdf1542850d66d8007" & _
                                        "d620e4050b5715dc83f4a921d36ce9ce" & _
                                        "47d0d13c5d85f2b0ff8318d2877eec2f" & _
                                        "63b931bd47417a81a538327af927da3e"
        End Select
    Else
        Data_Bytes = Recoding_Text(CStr(Source_Data), Type_ASCII)
        GoSub GS¦Hashing_Data
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````'
GS¦Hashing_Array:

    Data_Bytes = Conversion_ArrayToBytes(Source_Data)
    If UBound(Data_Bytes) = -1 Then Exit Function
    If IsArray(Data_Bytes) Then GoSub GS¦Hashing_Data
    
    Return
'`````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````'
GS¦Hashing_File:

    File_Path = Source_Data
    Get_HashSumm_Data = Conversion_FileToBytes(File_Path, HS_Algorithm)
    If IsEmpty(Get_HashSumm_Data) Then
        Get_HashSumm_Data = False
    Else
        Data_Bytes = Get_HashSumm_Data: GoSub GS¦Hashing_Data
    End If

    Return
'``````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦Hashing_Folder:

    If Right$(Source_Data, 1) <> "\" Then Source_Data = Source_Data & "\"
    ReDim Vector_HashSumm(1 To 10000): Inx = 0: Flag_SuccessfulHashSummCalculation = True
    File_Path = Dir(Source_Data & "*.*"): If File_Path = "" Then Exit Function

    Do While File_Path <> ""
        Get_HashSumm_Data = Conversion_FileToBytes(Source_Data & File_Path, HS_Algorithm)
        If IsEmpty(Get_HashSumm_Data) Then
            Flag_SuccessfulHashSummCalculation = False
        Else
            Data_Bytes = Get_HashSumm_Data: GoSub GS¦Hashing_Data: Inx = Inx + 1
            If Inx > UBound(Vector_HashSumm) Then
                ReDim Preserve Vector_HashSumm(1 To (Inx - 1) * 4)
            End If
            Vector_HashSumm(Inx) = Get_HashSumm_Data
        End If
        File_Path = Dir
    Loop

    If Not Flag_SuccessfulHashSummCalculation Then
        Get_HashSumm_Data = False
    Else
        If Inx <> UBound(Vector_HashSumm) Then
            ReDim Preserve Vector_HashSumm(1 To Inx)
        End If

        If Not Inx = 1 Then
            Data_Bytes = Recoding_Text(Join(Vector_HashSumm, vbNullString))
            GoSub GS¦Hashing_Data
        End If
    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````'
GS¦Hashing_VBComponent:

    For Each Obj_VBComponent In Application.VBE.ActiveVBProject.VBComponents
        If Obj_VBComponent.Name = Source_Data Then
            Set Obj_VBProject = Application.VBE.ActiveVBProject: Exit For
        End If
    Next

    If Obj_VBProject Is Nothing Then
        For Each VBProject In Application.VBE.VBProjects
            For Each Obj_VBComponent In VBProject.VBComponents
                If Obj_VBComponent.Name = Source_Data Then
                    Set Obj_VBProject = VBProject: Exit For
                End If
            Next Obj_VBComponent
            If Not Obj_VBProject Is Nothing Then Exit For
        Next VBProject
    End If

    File_Path = FileSystem_Unloading_VBComponent(CStr(Source_Data), Obj_VBProject)
    Data_Bytes = Conversion_FileToBytes(File_Path, HS_Algorithm)
    Kill File_Path: GoSub GS¦Hashing_Data
    If Right$(File_Path, 3) = "frm" Then
        File_Path = Left(File_Path, InStrRev(File_Path, ".") - 1) & ".frx"
        If Dir(File_Path) <> "" Then Kill File_Path
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````'
GS¦Hashing_VBProject:

    If Application.VBE.ActiveVBProject.Name = Source_Data Then
         Set Obj_VBProject = Application.VBE.ActiveVBProject
    Else
        For Each VBProject In Application.VBE.VBProjects
            If VBProject.Name = Source_Data Then
                Set Obj_VBProject = VBProject: Exit For
            End If
        Next VBProject
    End If

    Count_VBComponent = Obj_VBProject.VBComponents.Count
    ReDim Vector_HashSumm(1 To Count_VBComponent): Inx = 0

    For Each Obj_VBComponent In Obj_VBProject.VBComponents
        File_Path = FileSystem_Unloading_VBComponent(CStr(Obj_VBComponent.Name), _
                                                          Obj_VBProject)
        Data_Bytes = Conversion_FileToBytes(File_Path, HS_Algorithm)
        Kill File_Path: GoSub GS¦Hashing_Data: Inx = Inx + 1
        Vector_HashSumm(Inx) = Get_HashSumm_Data
    Next Obj_VBComponent

    Data_Bytes = Recoding_Text(Join(Vector_HashSumm, vbNullString))
    GoSub GS¦Hashing_Data
    
    Return
'`````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NonExistent_Directory:

    Error_Message = "Файл или папка не существует: " & Source_Data
    Call Show_ErrorMessage_Immediate(Error_Message, _
                                    "Ошибка поиска нужного файла/директории", , True)
    Get_HashSumm_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````'
GS¦Exception_NullString:

    Error_Message = "Нет данных для обработки: Len String = 0"
    Call Show_ErrorMessage_Immediate(Error_Message, "Передана пустая строка")
    Get_HashSumm_Data = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'======================================================================'
Private Sub Init_VBD_Kit_Hashing()
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


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'================================================================================================================'
Private Function Hashing_WinAPI( _
                 ByRef Data_Bytes() As Byte, _
                 ByRef Hashing_Algorithm As String _
        ) As String
'----------------------------------------------------------------------------------------------------------------'
    
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Return_HashSumm As String, Error_Message As String, Position As Long, Inx_LB1 As Long
    Dim Ptr_Data As LongPtr, Ptr_Hash As LongPtr, Ptr_Alg As LongPtr, Size_Data As Long, Inx_UB1   As Long
    Dim Buffer_Hash() As Byte, Buffer_HashObject() As Byte, Length_Hash As Long, Length As Long, i As Long
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    
    '```````````````````````````````````````````````````````````````````````````````'
    Select Case Hashing_Algorithm
        Case "MD5", "SHA1", "SHA256", "SHA384", "SHA512"
        Case Else
            Error_Message = "Нет поддержки данного алгоритма: " & Hashing_Algorithm
            Call Show_ErrorMessage_Immediate(Error_Message, "WinAPI_Exception")
            Hashing_WinAPI = vbNullString: Exit Function
    End Select
    '```````````````````````````````````````````````````````````````````````````````'
    
    '````````````````````````````````'
    Inx_LB1 = LBound(Data_Bytes, 1)
    Inx_UB1 = UBound(Data_Bytes, 1)

    Size_Data = Inx_UB1 - Inx_LB1 + 1
    '````````````````````````````````'

    '````````````````````````````````````'
    Ptr_Data = VarPtr(Data_Bytes(Inx_LB1))
    '````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next

    If BCryptOpenAlgorithmProvider(Ptr_Alg, _
                                   StrPtr(Hashing_Algorithm & vbNullChar), 0, 0) Then
        Error_Message = "BCryptOpenAlgorithmProvider: Ошибка доступа к провайдеру: " _
                                                                    & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function

    End If
    '``````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    If BCryptGetProperty(Ptr_Alg, _
                         StrPtr("ObjectLength" & vbNullString), Length, LenB(Length), 0, 0) <> 0 Then
        Error_Message = "BCryptGetProperty: Ошибка получения длины объекта хэша: " _
                                                                  & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    ReDim Buffer_HashObject(0 To Length - 1)

    If BCryptGetProperty(Ptr_Alg, _
                         StrPtr("HashDigestLength" & vbNullChar), Length_Hash, LenB(Length_Hash), 0, 0) <> 0 Then
        Error_Message = "BCryptGetProperty: Ошибка получения итоговой длины хэша: " _
                                                                   & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    ReDim Buffer_Hash(0 To Length_Hash - 1)

    If BCryptCreateHash(Ptr_Alg, Ptr_Hash, Buffer_HashObject(0), Length, 0, 0, 0) <> 0 Then
        Error_Message = "BCryptCreateHash: Ошибка создания объекта хэша: " _
                                                          & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    If BCryptHashData(Ptr_Hash, ByVal Ptr_Data, Size_Data) <> 0 Then
        Error_Message = "BCryptHashData: Ошибка добавления данных для хеширования: " _
                                                                    & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    If BCryptFinishHash(Ptr_Hash, Buffer_Hash(0), Length_Hash, 0) <> 0 Then
        Error_Message = "BCryptFinishHash: Ошибка извлечения значения хэша: " _
                                                           & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, "DLL_Error")
        GoSub GS¦Clearing_Memory: Hashing_WinAPI = vbNullString: Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````'

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

    Hashing_WinAPI = LCase$(Return_HashSumm)

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````'

''`````````````````````````````````````````````````````````````````'
GS¦Clearing_Memory:

    If CBool(Ptr_Hash) Then BCryptDestroyHash Ptr_Hash
    If CBool(Ptr_Alg) Then BCryptCloseAlgorithmProvider Ptr_Alg, 0

    On Error GoTo 0

    Return
''`````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================'


'==========================================================================================='
Private Function Hashing_NET( _
                 ByRef Data_Bytes() As Byte, _
                 ByRef Hashing_Algorithm As String _
        ) As String
'-------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````'
    Dim Obj_SSC_HashProvider As Object, Hash_Bytes() As Byte
    Dim Return_Hash   As String, Error_Message     As String
    '```````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    Select Case Hashing_Algorithm
        Case "MD5"
            Set Obj_SSC_HashProvider = _
                CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
        Case "SHA1"
            Set Obj_SSC_HashProvider = _
                CreateObject("System.Security.Cryptography.SHA1Managed")
        Case "SHA256"
            Set Obj_SSC_HashProvider = _
                CreateObject("System.Security.Cryptography.SHA256Managed")
        Case "SHA384"
            Set Obj_SSC_HashProvider = _
                CreateObject("System.Security.Cryptography.SHA384Managed")
        Case "SHA512"
            Set Obj_SSC_HashProvider = _
                CreateObject("System.Security.Cryptography.SHA512Managed")
        Case Else
            Error_Message = "Нет поддержки данного алгоритма: " & Hashing_Algorithm
            Call Show_ErrorMessage_Immediate(Error_Message, "NET_Exception")
            Hashing_NET = vbNullString: Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    Hash_Bytes = Obj_SSC_HashProvider.ComputeHash_2((Data_Bytes))
    '```````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    With CreateObject("MSXML2.DOMDocument")
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = Hash_Bytes
        Return_Hash = Replace(.DocumentElement.Text, vbLf, "")
    End With
    '`````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Hashing_NET = LCase$(Return_Hash)
    '```````````````````````````````'

'-------------------------------------------------------------------------------------------'
End Function
'==========================================================================================='


'==================================================================================='
Private Function Hashing_NativeCode( _
                 ByRef Data_Bytes() As Byte, _
                 ByRef Hashing_Algorithm As String _
        ) As String
'-----------------------------------------------------------------------------------'
    
    '``````````````````````````'
    Dim Error_Message As String
    '``````````````````````````'
    
    '```````````````````````````````````````````````````````````````````````````````'
    Select Case Hashing_Algorithm
        Case "MD5":    Hashing_NativeCode = MD5_Calculate(Data_Bytes)
        Case "SHA1":   Hashing_NativeCode = SHA1_Calculate(Data_Bytes)
        Case "SHA256": Hashing_NativeCode = SHA256_Calculate(Data_Bytes)
        Case "SHA384": Hashing_NativeCode = SHA384_Calculate(Data_Bytes)
        Case "SHA512": Hashing_NativeCode = SHA512_Calculate(Data_Bytes)
        Case Else
            Error_Message = "Нет поддержки данного алгоритма: " & Hashing_Algorithm
            Call Show_ErrorMessage_Immediate(Error_Message, "NativeVBA_Exception")
            Hashing_NativeCode = vbNullString: Exit Function
    End Select
    '```````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------'
End Function
'==================================================================================='


'-----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-----------------------------------------------------------------------------------------------------------------------------'


'========================================================================================='
Private Function Check_NetFramework_IsLoaded() As Boolean
'-----------------------------------------------------------------------------------------'

    '``````````````````````````````````'
    Dim MsgBox_Result As VbMsgBoxResult
    '``````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    If NetFramework_IsLoaded Then
        Check_NetFramework_IsLoaded = True
    Else
        Check_NetFramework_IsLoaded = False

        MsgBox_Result = MsgBox("Компоненты .NET Framework не обнаружены. " & vbNewLine & _
                               "Открыть окно для подключения компонентов .NET? ", _
                                vbYesNo + vbInformation, ".NET Framework")

        Select Case MsgBox_Result
            Case vbYes: Call NetFramework_ActivateWindowWithComponents
        End Select
    End If
    '`````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'================================================================'
Private Function NetFramework_IsLoaded() As Boolean
'----------------------------------------------------------------'

    '````````````````````````````````````````````````````````````'
    On Error Resume Next

    With CreateObject("System.Security.Cryptography.SHA1Managed")
        If Err.Number = 0 Then
            NetFramework_IsLoaded = True
        Else
            NetFramework_IsLoaded = False
        End If
    End With

    On Error GoTo 0
    '````````````````````````````````````````````````````````````'

'----------------------------------------------------------------'
End Function
'================================================================'


'============================================================================================================='
Private Sub NetFramework_ActivateWindowWithComponents(): Shell "control appwiz.cpl,,2", vbNormalFocus: End Sub
'============================================================================================================='


'-----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-----------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================'
Private Function Get_SemanticDataType( _
                 ByRef Data As Variant _
        ) As MSOffice_Type_ContentFormat_HS
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

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_VarType
        Case vbBoolean:              Get_SemanticDataType = MSDI_Bool_HS:        Exit Function
        Case vbEmpty:                Get_SemanticDataType = MSDI_Empty_HS:       Exit Function
        Case vbNull:                 Get_SemanticDataType = MSDI_Null_HS:        Exit Function
        Case vbError, vbObjectError: Get_SemanticDataType = MSDI_Undefined_HS:   Exit Function
        Case vbUserDefinedType:      Get_SemanticDataType = MSDI_UserDefType_HS: Exit Function
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If (Data_VarType And vbArray) And (Not Data_TypeName = "Range") Then
        Get_SemanticDataType = MSDI_Array_HS: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_TypeName
        Case "String"
            If Len(Data) = 0 Then Get_SemanticDataType = MSDI_NullString_HS: Exit Function
            If IsNumeric(Data) Then Get_SemanticDataType = MSDI_String_HS:   Exit Function
            Data_Type = Get_Directory_Type(CStr(Data))

            Select Case Data_Type
                Case DirType_File:     Get_SemanticDataType = MSDI_File_HS:                  Exit Function
                Case DirType_Folder:   Get_SemanticDataType = MSDI_Folder_HS:                Exit Function
                Case DirType_NotFound: Get_SemanticDataType = MSDI_NonExistent_Directory_HS: Exit Function
                Case DirType_Invalid:
                    On Error Resume Next
                    VBProject_Name = Application.VBE.ActiveVBProject.Name

                    If Not Err.Number = 0 Then
                        Error_Message = "Невозможно проверить тип данных (Компонент это или VB-Проект)! " & _
                                        "Тип данных будет определен как строка (String)!"
                        Call Show_ErrorMessage_Immediate(Error_Message, "Проблема идентификации типов")
                        Get_SemanticDataType = MSDI_String_HS: On Error GoTo 0: Exit Function
                    End If

                    If VBProject_Name = Data Then
                        Get_SemanticDataType = MSDI_VBProject_HS
                        On Error GoTo 0: Exit Function
                    End If

                    Err.Clear
                    VBComponent_Name = Application.VBE.ActiveVBProject.VBComponents.Item(Data).Name
                    
                    If Len(VBComponent_Name) > 0 And Err.Number = 0 Then
                        Get_SemanticDataType = MSDI_VBComponent_HS
                        On Error GoTo 0: Exit Function
                    End If

                    For Each Obj_VBProjects In Application.VBE.VBProjects
                        If Obj_VBProjects.Name = Data Then
                            Get_SemanticDataType = MSDI_VBProject_HS:   On Error GoTo 0: Exit Function
                        End If

                        Err.Clear: Set Obj_VBComponent = Obj_VBProjects.VBComponents.Item(Data)
                        If (Err.Number = 0) And (Not Obj_VBComponent Is Nothing) Then
                            If Not Len(Obj_VBComponent.Name) = 0 Then
                                Set Obj_VBComponent = Nothing
                                Get_SemanticDataType = MSDI_VBComponent_HS: On Error GoTo 0: Exit Function
                            End If
                        End If
                        Set Obj_VBComponent = Nothing
                    Next Obj_VBProjects

                    Get_SemanticDataType = MSDI_String_HS: On Error GoTo 0: Exit Function
            End Select

        Case "Byte":       Get_SemanticDataType = MSDI_Int8_HS:             Exit Function
        Case "Integer":    Get_SemanticDataType = MSDI_Int16_HS:            Exit Function
        Case "Long":       Get_SemanticDataType = MSDI_Int32_HS:            Exit Function
        Case "LongLong":   Get_SemanticDataType = MSDI_Int64_HS:            Exit Function
        Case "Single":     Get_SemanticDataType = MSDI_FloatPoint32_HS:     Exit Function
        Case "Currency":   Get_SemanticDataType = MSDI_FloatPoint64Cur_HS:  Exit Function
        Case "Double":     Get_SemanticDataType = MSDI_FloatPoint64Dbl_HS:  Exit Function
        Case "Decimal":    Get_SemanticDataType = MSDI_FloatPoint112_HS:    Exit Function
        Case "Date":       Get_SemanticDataType = MSDI_Date_HS:             Exit Function
        Case "Range":      Get_SemanticDataType = MSDI_Range_HS:            Exit Function
        Case "Dictionary": Get_SemanticDataType = MSDI_Dictionary_HS:       Exit Function
        Case "Collection": Get_SemanticDataType = MSDI_Collection_HS:       Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````'
    If IsObject(Data) Then
        If Not Data Is Nothing Then
            Get_SemanticDataType = MSDI_Object_HS:  Exit Function
        Else
            Get_SemanticDataType = MSDI_Nothing_HS: Exit Function
        End If
    Else
        Get_SemanticDataType = MSDI_Undefined_HS:   Exit Function
    End If
    '````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


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


'=============================================================='
Private Function FileSystem_Unloading_VBComponent( _
                 ByRef Name_Component As String, _
                 ByRef VBProject As Object _
        ) As String
'--------------------------------------------------------------'
    
    '````````````````````````````````````````````````````'
    Dim FullName_Component As String, File_Path As String
    Dim Obj_VBComponent    As Object
    '````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Set Obj_VBComponent = VBProject.VBComponents(Name_Component)

    Select Case Obj_VBComponent.Type
        Case 1:   FullName_Component = Name_Component & ".bas"
        Case 2:   FullName_Component = Name_Component & ".cls"
        Case 3:   FullName_Component = Name_Component & ".frm"
        Case 100: FullName_Component = Name_Component & ".cls"
    End Select
    '``````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    File_Path = Environ$("TEMP") & "\" & FullName_Component
    VBProject.VBComponents(Name_Component).Export File_Path
    '``````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    FileSystem_Unloading_VBComponent = File_Path
    '```````````````````````````````````````````'

'--------------------------------------------------------------'
End Function
'=============================================================='

    
'=========================================================================================================='
Private Function Conversion_FileToBytes( _
                 ByRef File_Path As String, _
                 Optional ByRef HS_Algorithm As String _
        ) As Variant
'----------------------------------------------------------------------------------------------------------'
    
    '```````````````````````````````````````````````````'
    Dim Return_HashSumm As String
    Dim Size_File As Double, FF As Integer
    Dim Return_Bytes() As Byte, Data_Bytes()  As Byte
    Dim Handle_File As Long, Flag_Done        As Boolean
    Dim Buffer_Bytes() As Byte, Handled_Bytes As Long
    '```````````````````````````````````````````````````'
    Dim Total_HashSumm As String, Bytes_Read  As Double
    Dim Read_StartPosition As Long, Len_Hash  As Long
    Dim Error_Message As String, Num_ByteRead As Long
    '```````````````````````````````````````````````````'
    #If Win64 Then
        Const BlockSize_Hashing As Long = &H1000000
    #Else
        Const BlockSize_Hashing As Long = &H3FFFFF
    #End If
    '```````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````'
    Conversion_FileToBytes = Empty

    If File_IsOpen(File_Path) Then
        Error_Message = "Файл открыт и не может быть прочитан (" & File_Path & ")!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу")
        Exit Function
    End If

    Size_File = Get_FileSize(File_Path)

    If Size_File = -1 Then
        Error_Message = "Не удалось открыть файл. Возможно проблема с правами доступа (" & File_Path & ")!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Нет доступа к файлу")
        Exit Function
    End If

    If Size_File = 0 Then
        Error_Message = "Нет данных для хеширования! Размер файла равен 0 (" & File_Path & ")!"
        Call Show_ErrorMessage_Immediate(Error_Message, "Отсутствуют данные для обработки")
        Exit Function
    End If

    Select Case Size_File
        Case Is > 2147483648#
            GoSub GS¦ReadingFile_More_2GB
        Case Is > BlockSize_Hashing * 64
            #If Win64 Then
                GoSub GS¦ReadingFile_Less_2GB
            #Else
                GoSub GS¦ReadingFile_More_2GB
            #End If
        Case Else
            GoSub GS¦ReadingFile_Less_2GB
    End Select

    Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````````````````````'

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

'``````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ReadingFile_More_2GB:

    Flag_Done = True: Read_StartPosition = 1

    Handle_File = CreateFile(File_Path, GENERIC_WRITE Or GENERIC_READ, _
                             FILE_SHARE_DEFAULT, 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_READONLY, 0&)

    If Handle_File = -1 Then
        Error_Message = "Файл в данный момент открыт: " _
                                       & Err.LastDllError
        Call Show_ErrorMessage_Immediate(Error_Message, File_Path)
        Conversion_FileToBytes = vbNullString: Exit Function
    End If

    Total_HashSumm = Space(Size_File / BlockSize_Hashing * 256)

    While Flag_Done
        If Size_File > Bytes_Read + BlockSize_Hashing Then
            Handled_Bytes = BlockSize_Hashing
        Else
            Handled_Bytes = Size_File - Bytes_Read: Flag_Done = False
        End If

        ReDim Buffer_Bytes(0 To Handled_Bytes - 1) As Byte

        ReadFile Handle_File, Buffer_Bytes(0), Handled_Bytes, Num_ByteRead, 0
        Data_Bytes = Buffer_Bytes: Return_HashSumm = Hashing_WinAPI(Data_Bytes, HS_Algorithm)
        Len_Hash = Len(Return_HashSumm)

        Mid$(Total_HashSumm, Read_StartPosition, Len_Hash) = Return_HashSumm

        Read_StartPosition = Read_StartPosition + Len_Hash
        Bytes_Read = Bytes_Read + Handled_Bytes
    Wend

    CloseHandle Handle_File
    Conversion_FileToBytes = Recoding_Text(RTrim$(Total_HashSumm))

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Throw_OutOfMemory:

    Error_Message = "Обновите MS Office до x64 для расширения максимального объёма памяти!"
    Call Show_ErrorMessage_Immediate(Error_Message, "Недостаточно памяти для считывания файла")

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================='


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


'======================================================================================'
Private Function Get_FileSize( _
                 ByVal FileSystem_FilePath As String, _
                 Optional Format_Size As FileSystem_Format_SizeFile = SizeFormat_Byte _
        ) As Double
'--------------------------------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Handle_File As LongPtr, result As Double
    Dim File_Size   As WinAPI_LargeInteger
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

'--------------------------------------------------------------------------------------'
End Function
'======================================================================================'


'============================================================================='
Private Function Conversion_ArrayToBytes( _
                 ByRef Data As Variant _
        ) As Variant
'-----------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Return_Bytes() As Byte, Error_Message As String
    Dim Inx_UB1 As Long, tmp_String As String, tmp_Data
    '``````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    If TypeName(Data) = "Range" Then
        Select Case Glb_MSOffice_Type_Application
            Case Type_MSOffice_Excel
                tmp_Data = Data.Value
                If Data.Count = 1 Then
                    ReDim tmp_Data(0) As String: tmp_Data(0) = Data.Value
                End If
            Case Type_MSOffice_Word
                ReDim tmp_Data(0) As String: tmp_Data(0) = Data.Text
        End Select
    Else
        tmp_Data = Data
    End If
    '`````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    On Error Resume Next: Inx_UB1 = UBound(tmp_Data)
    '```````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If Err.Number = 0 Then
        If TypeName(tmp_Data) = "Byte()" Then
            Conversion_ArrayToBytes = Data
            Exit Function
        End If

        If TypeName(tmp_Data) = "Object()" Then
            Error_Message = "Тип данных не поддерживается: " _
                                        & "Object()"
            Call Show_ErrorMessage_Immediate(Error_Message, _
                                            "Несоответствие типов")
            Conversion_ArrayToBytes = vbNullString
            Exit Function
        Else
            tmp_String = Conversion_ArrayToString(tmp_Data)

            If Len(tmp_String) = 0 Then
                Error_Message = "Массив пуст и не имеет данных!"
                Call Show_ErrorMessage_Immediate(Error_Message, _
                                                "Нет данных для хеширования")
                Conversion_ArrayToBytes = vbNullString
                Exit Function
            End If

            Return_Bytes = Recoding_Text(tmp_String)
            Conversion_ArrayToBytes = Return_Bytes
            Exit Function
        End If
    Else
        Error_Message = "Массив не инициализирован!"
        Call Show_ErrorMessage_Immediate(Error_Message, _
                                         "Нет данных для хеширования")
        Conversion_ArrayToBytes = vbNullString
    End If
    '`````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------'
End Function
'============================================================================='


'========================================================================================='
Private Function Conversion_ArrayToString( _
                 ByRef Source_Array As Variant _
        ) As String
'-----------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````'
    Dim FullLen_String As Long
    Dim Position       As Long, Len_String   As Long, Item_Array As Variant
    Dim Return_String  As String, tmp_String As String, Len_Cell As Long
    '``````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    FullLen_String = 0: Return_String = String$(1, vbNullChar)
    '````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next

    For Each Item_Array In Source_Array
        Len_Cell = Len(Item_Array)
        If Err.Number <> 0 Then Item_Array = CStr(Item_Array): Len_Cell = Len(Item_Array)

        If Len_Cell > 0 Then
            Len_String = Len(Return_String): FullLen_String = Position + Len_Cell

            If FullLen_String > Len_String Then Return_String = _
                                                Return_String & _
                                                String$(Len_String + Len_Cell, vbNullChar)

            Mid$(Return_String, Position + 1) = Item_Array: Position = FullLen_String
        End If
    Next Item_Array

    On Error GoTo 0
    '`````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    If FullLen_String = 0 Then Conversion_ArrayToString = vbNullString: Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    Conversion_ArrayToString = Left$(Return_String, FullLen_String)
    '``````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'================================================================================'
Private Function Recoding_Text( _
                 ByRef Source_Text As String, _
                 Optional ByVal Encoding As MSOffice_Type_Encoding = Type_UTF_8 _
        ) As Byte()
'--------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````'
    Dim Byte_Result() As Byte, Ptr_Source As LongPtr
    Dim i  As Long, J As Long, Char_Set   As String, Length As Long
    '``````````````````````````````````````````````````````````````'
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


'-----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-----------------------------------------------------------------------------------------------------------------------------'


'==========================================================================================='
Private Function Show_ErrorMessage_Immediate( _
                 ByRef Error_Message As String, _
                 Optional ByRef Header_Message As Variant = vbNullChar, _
                 Optional ByRef WinAPI_DllName As String = vbNullString, _
                 Optional ByVal Show_Immediate As Boolean = False _
        ) As String
'-------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim RT_Error_Number As Long, RT_Error_Description As String
    Dim Len_Description As Long, Len_Message As Long, Len_Borders As Long
    Dim WB_Name  As String, Debug_Borders    As String
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    RT_Error_Number = Err.Number: RT_Error_Description = Err.Description
    RT_Error_Description = Replace(RT_Error_Description, ChrW$(10), vbNullChar)
    '``````````````````````````````````````````````````````````````````````````'

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


'=========================================================================='
Private Function MD5_Calculate( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'--------------------------------------------------------------------------'
    Call MD5_Start
    Call MD5_64Split(Len(CStr(StrConv(Data_Bytes, vbUnicode))), Data_Bytes)
    Call MD5_Finish: MD5_Calculate = MD5_Value
'--------------------------------------------------------------------------'
End Function
'=========================================================================='


'==============================================================='
Private Function SHA1_Calculate( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'---------------------------------------------------------------'
    SHA1_Calculate = SHA1_HexDefault(Data_Bytes)
'---------------------------------------------------------------'
End Function
'==============================================================='


'==============================================================='
Private Function SHA256_Calculate( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'---------------------------------------------------------------'
    SHA256_Calculate = SHA256_Start(Data_Bytes)
'---------------------------------------------------------------'
End Function
'==============================================================='


'==============================================================='
Private Function SHA384_Calculate( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'---------------------------------------------------------------'
    SHA384_Calculate = SHA512_CryptoSha512Text(384, Data_Bytes)
'---------------------------------------------------------------'
End Function
'==============================================================='


'==============================================================='
Private Function SHA512_Calculate( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'---------------------------------------------------------------'
    SHA512_Calculate = SHA512_CryptoSha512Text(512, Data_Bytes)
'---------------------------------------------------------------'
End Function
'==============================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================='
Private Sub MD5_Start()
'-------------------------------------------------------------'
    MD5_lngTrack = 0
    MD5_ArrLongConversion(1) = MD5_LongConversion(1732584193#)
    MD5_ArrLongConversion(2) = MD5_LongConversion(4023233417#)
    MD5_ArrLongConversion(3) = MD5_LongConversion(2562383102#)
    MD5_ArrLongConversion(4) = MD5_LongConversion(271733878#)
'-------------------------------------------------------------'
End Sub
'============================================================='


'==================================================================='
Private Function MD5_Round( _
                 ByRef strRound As String, _
                 ByRef A As Long, _
                 ByRef B As Long, _
                 ByRef C As Long, _
                 ByRef D As Long, _
                 ByRef x As Long, _
                 ByRef S As Long, _
                 ByRef AC As Long _
        ) As Long
'-------------------------------------------------------------------'
    Select Case strRound
        Case Is = "FF"
            A = MD5_LongAdd4(A, (B And C) Or (Not (B) And D), x, AC)
            A = MD5_Rotate(A, S)
            A = MD5_LongAdd(A, B)
        Case Is = "GG"
            A = MD5_LongAdd4(A, (B And D) Or (C And Not (D)), x, AC)
            A = MD5_Rotate(A, S)
            A = MD5_LongAdd(A, B)
        Case Is = "HH"
            A = MD5_LongAdd4(A, B Xor C Xor D, x, AC)
            A = MD5_Rotate(A, S)
            A = MD5_LongAdd(A, B)
        Case Is = "II"
            A = MD5_LongAdd4(A, C Xor (B Or Not (D)), x, AC)
            A = MD5_Rotate(A, S)
            A = MD5_LongAdd(A, B)
    End Select
'-------------------------------------------------------------------'
End Function
'==================================================================='


'====================================================================================================='
Private Function MD5_Rotate( _
                 ByRef lngValue As Long, _
                 ByRef lngBits As Long _
        ) As Long
'-----------------------------------------------------------------------------------------------------'
    Dim lngSign As Long, lngI As Long

    lngBits = (lngBits Mod 32)
    If lngBits = 0 Then MD5_Rotate = lngValue: Exit Function

    For lngI = 1 To lngBits
        lngSign = lngValue And &HC0000000
        lngValue = (lngValue And &H3FFFFFFF) * 2
        lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next

    MD5_Rotate = lngValue
'-----------------------------------------------------------------------------------------------------'
End Function
'====================================================================================================='


'================================================================================'
Private Function MD5_64Split( _
                 ByRef lngLength As Long, _
                 ByRef bytBuffer() As Byte _
        ) As String
'--------------------------------------------------------------------------------'
    Dim lngBytesTotal As Long, lngBytesToAdd As Long
    Dim intLoop As Long, intLoop2 As Long, lngTrace As Long
    Dim intInnerLoop As Long, intLoop3 As Long

    lngBytesTotal = MD5_lngTrack Mod 64
    lngBytesToAdd = 64 - lngBytesTotal
    MD5_lngTrack = (MD5_lngTrack + lngLength)

    If lngLength >= lngBytesToAdd Then
        For intLoop = 0 To lngBytesToAdd - 1
            MD5_ArrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
        Next intLoop

        MD5_Conversion MD5_ArrSplit64

        lngTrace = (lngLength) Mod 64

        For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64
            For intInnerLoop = 0 To 63
                MD5_ArrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
            Next intInnerLoop

            MD5_Conversion MD5_ArrSplit64

        Next intLoop2

        lngBytesTotal = 0
    Else
        intLoop2 = 0
    End If

    For intLoop3 = 0 To lngLength - intLoop2 - 1
        MD5_ArrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)
    Next intLoop3
'--------------------------------------------------------------------------------'
End Function
'================================================================================'


'================================================================================'
Private Sub MD5_Conversion( _
            ByRef bytBuffer() As Byte _
        )
'--------------------------------------------------------------------------------'
    Dim x(16) As Long, A As Long
    Dim B As Long, C As Long
    Dim D As Long

    A = MD5_ArrLongConversion(1)
    B = MD5_ArrLongConversion(2)
    C = MD5_ArrLongConversion(3)
    D = MD5_ArrLongConversion(4)

    MD5_Decode 64, x, bytBuffer

    MD5_Round "FF", A, B, C, D, x(0), MD5_S11, -680876936
    MD5_Round "FF", D, A, B, C, x(1), MD5_S12, -389564586
    MD5_Round "FF", C, D, A, B, x(2), MD5_S13, 606105819
    MD5_Round "FF", B, C, D, A, x(3), MD5_S14, -1044525330
    MD5_Round "FF", A, B, C, D, x(4), MD5_S11, -176418897
    MD5_Round "FF", D, A, B, C, x(5), MD5_S12, 1200080426
    MD5_Round "FF", C, D, A, B, x(6), MD5_S13, -1473231341
    MD5_Round "FF", B, C, D, A, x(7), MD5_S14, -45705983
    MD5_Round "FF", A, B, C, D, x(8), MD5_S11, 1770035416
    MD5_Round "FF", D, A, B, C, x(9), MD5_S12, -1958414417
    MD5_Round "FF", C, D, A, B, x(10), MD5_S13, -42063
    MD5_Round "FF", B, C, D, A, x(11), MD5_S14, -1990404162
    MD5_Round "FF", A, B, C, D, x(12), MD5_S11, 1804603682
    MD5_Round "FF", D, A, B, C, x(13), MD5_S12, -40341101
    MD5_Round "FF", C, D, A, B, x(14), MD5_S13, -1502002290
    MD5_Round "FF", B, C, D, A, x(15), MD5_S14, 1236535329

    MD5_Round "GG", A, B, C, D, x(1), MD5_S21, -165796510
    MD5_Round "GG", D, A, B, C, x(6), MD5_S22, -1069501632
    MD5_Round "GG", C, D, A, B, x(11), MD5_S23, 643717713
    MD5_Round "GG", B, C, D, A, x(0), MD5_S24, -373897302
    MD5_Round "GG", A, B, C, D, x(5), MD5_S21, -701558691
    MD5_Round "GG", D, A, B, C, x(10), MD5_S22, 38016083
    MD5_Round "GG", C, D, A, B, x(15), MD5_S23, -660478335
    MD5_Round "GG", B, C, D, A, x(4), MD5_S24, -405537848
    MD5_Round "GG", A, B, C, D, x(9), MD5_S21, 568446438
    MD5_Round "GG", D, A, B, C, x(14), MD5_S22, -1019803690
    MD5_Round "GG", C, D, A, B, x(3), MD5_S23, -187363961
    MD5_Round "GG", B, C, D, A, x(8), MD5_S24, 1163531501
    MD5_Round "GG", A, B, C, D, x(13), MD5_S21, -1444681467
    MD5_Round "GG", D, A, B, C, x(2), MD5_S22, -51403784
    MD5_Round "GG", C, D, A, B, x(7), MD5_S23, 1735328473
    MD5_Round "GG", B, C, D, A, x(12), MD5_S24, -1926607734

    MD5_Round "HH", A, B, C, D, x(5), MD5_S31, -378558
    MD5_Round "HH", D, A, B, C, x(8), MD5_S32, -2022574463
    MD5_Round "HH", C, D, A, B, x(11), MD5_S33, 1839030562
    MD5_Round "HH", B, C, D, A, x(14), MD5_S34, -35309556
    MD5_Round "HH", A, B, C, D, x(1), MD5_S31, -1530992060
    MD5_Round "HH", D, A, B, C, x(4), MD5_S32, 1272893353
    MD5_Round "HH", C, D, A, B, x(7), MD5_S33, -155497632
    MD5_Round "HH", B, C, D, A, x(10), MD5_S34, -1094730640
    MD5_Round "HH", A, B, C, D, x(13), MD5_S31, 681279174
    MD5_Round "HH", D, A, B, C, x(0), MD5_S32, -358537222
    MD5_Round "HH", C, D, A, B, x(3), MD5_S33, -722521979
    MD5_Round "HH", B, C, D, A, x(6), MD5_S34, 76029189
    MD5_Round "HH", A, B, C, D, x(9), MD5_S31, -640364487
    MD5_Round "HH", D, A, B, C, x(12), MD5_S32, -421815835
    MD5_Round "HH", C, D, A, B, x(15), MD5_S33, 530742520
    MD5_Round "HH", B, C, D, A, x(2), MD5_S34, -995338651

    MD5_Round "II", A, B, C, D, x(0), MD5_S41, -198630844
    MD5_Round "II", D, A, B, C, x(7), MD5_S42, 1126891415
    MD5_Round "II", C, D, A, B, x(14), MD5_S43, -1416354905
    MD5_Round "II", B, C, D, A, x(5), MD5_S44, -57434055
    MD5_Round "II", A, B, C, D, x(12), MD5_S41, 1700485571
    MD5_Round "II", D, A, B, C, x(3), MD5_S42, -1894986606
    MD5_Round "II", C, D, A, B, x(10), MD5_S43, -1051523
    MD5_Round "II", B, C, D, A, x(1), MD5_S44, -2054922799
    MD5_Round "II", A, B, C, D, x(8), MD5_S41, 1873313359
    MD5_Round "II", D, A, B, C, x(15), MD5_S42, -30611744
    MD5_Round "II", C, D, A, B, x(6), MD5_S43, -1560198380
    MD5_Round "II", B, C, D, A, x(13), MD5_S44, 1309151649
    MD5_Round "II", A, B, C, D, x(4), MD5_S41, -145523070
    MD5_Round "II", D, A, B, C, x(11), MD5_S42, -1120210379
    MD5_Round "II", C, D, A, B, x(2), MD5_S43, 718787259
    MD5_Round "II", B, C, D, A, x(9), MD5_S44, -343485551

    MD5_ArrLongConversion(1) = MD5_LongAdd(MD5_ArrLongConversion(1), A)
    MD5_ArrLongConversion(2) = MD5_LongAdd(MD5_ArrLongConversion(2), B)
    MD5_ArrLongConversion(3) = MD5_LongAdd(MD5_ArrLongConversion(3), C)
    MD5_ArrLongConversion(4) = MD5_LongAdd(MD5_ArrLongConversion(4), D)
'--------------------------------------------------------------------------------'
End Sub
'================================================================================'


'===================================================================================================================='
Private Function MD5_LongAdd( _
                 ByRef lngVal1 As Long, _
                 ByRef lngVal2 As Long _
        ) As Long
'--------------------------------------------------------------------------------------------------------------------'
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&

    MD5_LongAdd = MD5_LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
'--------------------------------------------------------------------------------------------------------------------'
End Function
'===================================================================================================================='


'===================================================================================================================='
Private Function MD5_LongAdd4( _
                 ByRef lngVal1 As Long, _
                 ByRef lngVal2 As Long, _
                 ByRef lngVal3 As Long, _
                 ByRef lngVal4 As Long _
        ) As Long
'--------------------------------------------------------------------------------------------------------------------'
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + _
                   ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    MD5_LongAdd4 = MD5_LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
'--------------------------------------------------------------------------------------------------------------------'
End Function
'===================================================================================================================='


'===================================================================================================================='
Private Sub MD5_Decode( _
            ByRef intLength As Integer, _
            ByRef lngOutBuffer() As Long, _
            ByRef bytInBuffer() As Byte _
        )
'--------------------------------------------------------------------------------------------------------------------'
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double

    intDblIndex = 0

    For intByteIndex = 0 To intLength - 1 Step 4
        dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) _
                                                               * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
        lngOutBuffer(intDblIndex) = MD5_LongConversion(dblSum)
        intDblIndex = (intDblIndex + 1)
    Next intByteIndex
'--------------------------------------------------------------------------------------------------------------------'
End Sub
'===================================================================================================================='


'=============================================================='
Private Function MD5_LongConversion( _
                 ByRef dblValue As Double _
        ) As Long
'--------------------------------------------------------------'
    If dblValue < 0 Or dblValue >= MD5_OFFSET_4 Then Error 6

    If dblValue <= MD5_MAXINT_4 Then
        MD5_LongConversion = dblValue
    Else
        MD5_LongConversion = dblValue - MD5_OFFSET_4
    End If
'--------------------------------------------------------------'
End Function
'=============================================================='


'==================================================================='
Private Sub MD5_Finish()
'-------------------------------------------------------------------'
    Dim dblBits As Double
    Dim arrPadding(72) As Byte
    Dim lngBytesBuffered As Long

    arrPadding(0) = &H80

    dblBits = MD5_lngTrack * 8

    lngBytesBuffered = MD5_lngTrack Mod 64

    If lngBytesBuffered <= 56 Then
        MD5_64Split (56 - lngBytesBuffered), arrPadding
    Else
        MD5_64Split (120 - MD5_lngTrack), arrPadding
    End If

    arrPadding(0) = MD5_LongConversion(dblBits) And &HFF&
    arrPadding(1) = MD5_LongConversion(dblBits) \ 256 And &HFF&
    arrPadding(2) = MD5_LongConversion(dblBits) \ 65536 And &HFF&
    arrPadding(3) = MD5_LongConversion(dblBits) \ 16777216 And &HFF&
    arrPadding(4) = 0
    arrPadding(5) = 0
    arrPadding(6) = 0
    arrPadding(7) = 0

    MD5_64Split 8, arrPadding
'-------------------------------------------------------------------'
End Sub
'==================================================================='


'=============================================================='
Private Function MD5_StringChange( _
                 ByRef lngnum As Long _
        ) As String
'--------------------------------------------------------------'
    Dim bytA As Byte
    Dim bytB As Byte
    Dim bytC As Byte
    Dim bytD As Byte

    bytA = lngnum And &HFF&
    If bytA < 16 Then
        MD5_StringChange = "0" & Hex(bytA)
    Else
        MD5_StringChange = Hex(bytA)
    End If

    bytB = (lngnum And &HFF00&) \ 256
    If bytB < 16 Then
        MD5_StringChange = MD5_StringChange & "0" & Hex(bytB)
    Else
        MD5_StringChange = MD5_StringChange & Hex(bytB)
    End If

    bytC = (lngnum And &HFF0000) \ 65536
    If bytC < 16 Then
        MD5_StringChange = MD5_StringChange & "0" & Hex(bytC)
    Else
        MD5_StringChange = MD5_StringChange & Hex(bytC)
    End If

    If lngnum < 0 Then
        bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
    Else
        bytD = (lngnum And &HFF000000) \ 16777216
    End If

    If bytD < 16 Then
        MD5_StringChange = MD5_StringChange & "0" & Hex(bytD)
    Else
        MD5_StringChange = MD5_StringChange & Hex(bytD)
    End If
'--------------------------------------------------------------'
End Function
'=============================================================='


'==================================================================='
Private Function MD5_Value() As String
'-------------------------------------------------------------------'
    MD5_Value = LCase(MD5_StringChange(MD5_ArrLongConversion(1)) & _
                      MD5_StringChange(MD5_ArrLongConversion(2)) & _
                      MD5_StringChange(MD5_ArrLongConversion(3)) & _
                      MD5_StringChange(MD5_ArrLongConversion(4)))
'-------------------------------------------------------------------'
End Function
'==================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


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


'==============================================================================================================================================='
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
'-----------------------------------------------------------------------------------------------------------------------------------------------'
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
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key3), ((B And C) Or (B And D) Or (C And D)))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i
         For i = 60 To 79
             T = SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32Add(SHA1_U32RotateLeft5(A), E), W(i)), Key4), (B Xor C Xor D))
             E = D: D = C: C = SHA1_U32RotateLeft30(B): B = A: A = T
         Next i

         H1 = SHA1_U32Add(H1, A): H2 = SHA1_U32Add(H2, B): H3 = SHA1_U32Add(H3, C): H4 = SHA1_U32Add(H4, D): H5 = SHA1_U32Add(H5, E)
     Wend
'-----------------------------------------------------------------------------------------------------------------------------------------------'
End Sub
'==============================================================================================================================================='


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


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'==================================================================================================================='
Private Function SHA256_Start( _
                 ByRef Data_Bytes() As Byte _
        ) As String
'-------------------------------------------------------------------------------------------------------------------'
    Dim Hash(7) As Long
    Dim M()     As Long
    Dim W(63)   As Long
    Dim A       As Long
    Dim B       As Long
    Dim C       As Long
    Dim D       As Long
    Dim E       As Long
    Dim F       As Long
    Dim G       As Long
    Dim H       As Long
    Dim i       As Long
    Dim J       As Long
    Dim T1      As Long
    Dim T2      As Long

    Dim Len_Data    As Long
    Dim Data_String As String

    Static SHA256_Initialized As Boolean

    If Not SHA256_Initialized Then
        SHA256_m_lOnBits(0) = 1
        SHA256_m_lOnBits(1) = 3
        SHA256_m_lOnBits(2) = 7
        SHA256_m_lOnBits(3) = 15
        SHA256_m_lOnBits(4) = 31
        SHA256_m_lOnBits(5) = 63
        SHA256_m_lOnBits(6) = 127
        SHA256_m_lOnBits(7) = 255
        SHA256_m_lOnBits(8) = 511
        SHA256_m_lOnBits(9) = 1023
        SHA256_m_lOnBits(10) = 2047
        SHA256_m_lOnBits(11) = 4095
        SHA256_m_lOnBits(12) = 8191
        SHA256_m_lOnBits(13) = 16383
        SHA256_m_lOnBits(14) = 32767
        SHA256_m_lOnBits(15) = 65535
        SHA256_m_lOnBits(16) = 131071
        SHA256_m_lOnBits(17) = 262143
        SHA256_m_lOnBits(18) = 524287
        SHA256_m_lOnBits(19) = 1048575
        SHA256_m_lOnBits(20) = 2097151
        SHA256_m_lOnBits(21) = 4194303
        SHA256_m_lOnBits(22) = 8388607
        SHA256_m_lOnBits(23) = 16777215
        SHA256_m_lOnBits(24) = 33554431
        SHA256_m_lOnBits(25) = 67108863
        SHA256_m_lOnBits(26) = 134217727
        SHA256_m_lOnBits(27) = 268435455
        SHA256_m_lOnBits(28) = 536870911
        SHA256_m_lOnBits(29) = 1073741823
        SHA256_m_lOnBits(30) = 2147483647

        SHA256_m_l2Power(0) = 1
        SHA256_m_l2Power(1) = 2
        SHA256_m_l2Power(2) = 4
        SHA256_m_l2Power(3) = 8
        SHA256_m_l2Power(4) = 16
        SHA256_m_l2Power(5) = 32
        SHA256_m_l2Power(6) = 64
        SHA256_m_l2Power(7) = 128
        SHA256_m_l2Power(8) = 256
        SHA256_m_l2Power(9) = 512
        SHA256_m_l2Power(10) = 1024
        SHA256_m_l2Power(11) = 2048
        SHA256_m_l2Power(12) = 4096
        SHA256_m_l2Power(13) = 8192
        SHA256_m_l2Power(14) = 16384
        SHA256_m_l2Power(15) = 32768
        SHA256_m_l2Power(16) = 65536
        SHA256_m_l2Power(17) = 131072
        SHA256_m_l2Power(18) = 262144
        SHA256_m_l2Power(19) = 524288
        SHA256_m_l2Power(20) = 1048576
        SHA256_m_l2Power(21) = 2097152
        SHA256_m_l2Power(22) = 4194304
        SHA256_m_l2Power(23) = 8388608
        SHA256_m_l2Power(24) = 16777216
        SHA256_m_l2Power(25) = 33554432
        SHA256_m_l2Power(26) = 67108864
        SHA256_m_l2Power(27) = 134217728
        SHA256_m_l2Power(28) = 268435456
        SHA256_m_l2Power(29) = 536870912
        SHA256_m_l2Power(30) = 1073741824

        SHA256_K(0) = &H428A2F98
        SHA256_K(1) = &H71374491
        SHA256_K(2) = &HB5C0FBCF
        SHA256_K(3) = &HE9B5DBA5
        SHA256_K(4) = &H3956C25B
        SHA256_K(5) = &H59F111F1
        SHA256_K(6) = &H923F82A4
        SHA256_K(7) = &HAB1C5ED5
        SHA256_K(8) = &HD807AA98
        SHA256_K(9) = &H12835B01
        SHA256_K(10) = &H243185BE
        SHA256_K(11) = &H550C7DC3
        SHA256_K(12) = &H72BE5D74
        SHA256_K(13) = &H80DEB1FE
        SHA256_K(14) = &H9BDC06A7
        SHA256_K(15) = &HC19BF174
        SHA256_K(16) = &HE49B69C1
        SHA256_K(17) = &HEFBE4786
        SHA256_K(18) = &HFC19DC6
        SHA256_K(19) = &H240CA1CC
        SHA256_K(20) = &H2DE92C6F
        SHA256_K(21) = &H4A7484AA
        SHA256_K(22) = &H5CB0A9DC
        SHA256_K(23) = &H76F988DA
        SHA256_K(24) = &H983E5152
        SHA256_K(25) = &HA831C66D
        SHA256_K(26) = &HB00327C8
        SHA256_K(27) = &HBF597FC7
        SHA256_K(28) = &HC6E00BF3
        SHA256_K(29) = &HD5A79147
        SHA256_K(30) = &H6CA6351
        SHA256_K(31) = &H14292967
        SHA256_K(32) = &H27B70A85
        SHA256_K(33) = &H2E1B2138
        SHA256_K(34) = &H4D2C6DFC
        SHA256_K(35) = &H53380D13
        SHA256_K(36) = &H650A7354
        SHA256_K(37) = &H766A0ABB
        SHA256_K(38) = &H81C2C92E
        SHA256_K(39) = &H92722C85
        SHA256_K(40) = &HA2BFE8A1
        SHA256_K(41) = &HA81A664B
        SHA256_K(42) = &HC24B8B70
        SHA256_K(43) = &HC76C51A3
        SHA256_K(44) = &HD192E819
        SHA256_K(45) = &HD6990624
        SHA256_K(46) = &HF40E3585
        SHA256_K(47) = &H106AA070
        SHA256_K(48) = &H19A4C116
        SHA256_K(49) = &H1E376C08
        SHA256_K(50) = &H2748774C
        SHA256_K(51) = &H34B0BCB5
        SHA256_K(52) = &H391C0CB3
        SHA256_K(53) = &H4ED8AA4A
        SHA256_K(54) = &H5B9CCA4F
        SHA256_K(55) = &H682E6FF3
        SHA256_K(56) = &H748F82EE
        SHA256_K(57) = &H78A5636F
        SHA256_K(58) = &H84C87814
        SHA256_K(59) = &H8CC70208
        SHA256_K(60) = &H90BEFFFA
        SHA256_K(61) = &HA4506CEB
        SHA256_K(62) = &HBEF9A3F7
        SHA256_K(63) = &HC67178F2

        SHA256_Initialized = True
    End If
    
    Hash(0) = &H6A09E667
    Hash(1) = &HBB67AE85
    Hash(2) = &H3C6EF372
    Hash(3) = &HA54FF53A
    Hash(4) = &H510E527F
    Hash(5) = &H9B05688C
    Hash(6) = &H1F83D9AB
    Hash(7) = &H5BE0CD19

    Len_Data = UBound(Data_Bytes) - LBound(Data_Bytes) + 1
    Data_String = String$(Len_Data, vbNullChar)

    For i = LBound(Data_Bytes) To UBound(Data_Bytes)
        Mid$(Data_String, i + 1) = ChrW(Data_Bytes(i))
    Next i

    M = SHA256_ConvertToWordArray(Data_String)

    For i = 0 To UBound(M) Step 16
        A = Hash(0)
        B = Hash(1)
        C = Hash(2)
        D = Hash(3)
        E = Hash(4)
        F = Hash(5)
        G = Hash(6)
        H = Hash(7)

        For J = 0 To 63
            If J < 16 Then
                W(J) = M(J + i)
            Else
                W(J) = SHA256_AddUnsigned(SHA256_AddUnsigned(SHA256_AddUnsigned(SHA256_Gamma1(W(J - 2)), _
                    W(J - 7)), SHA256_Gamma0(W(J - 15))), W(J - 16))
            End If

            T1 = SHA256_AddUnsigned(SHA256_AddUnsigned(SHA256_AddUnsigned(SHA256_AddUnsigned(H, SHA256_Sigma1(E)), _
                SHA256_Ch(E, F, G)), SHA256_K(J)), W(J))
            T2 = SHA256_AddUnsigned(SHA256_Sigma0(A), SHA256_Maj(A, B, C))

            H = G
            G = F
            F = E
            E = SHA256_AddUnsigned(D, T1)
            D = C
            C = B
            B = A
            A = SHA256_AddUnsigned(T1, T2)
        Next

        Hash(0) = SHA256_AddUnsigned(A, Hash(0))
        Hash(1) = SHA256_AddUnsigned(B, Hash(1))
        Hash(2) = SHA256_AddUnsigned(C, Hash(2))
        Hash(3) = SHA256_AddUnsigned(D, Hash(3))
        Hash(4) = SHA256_AddUnsigned(E, Hash(4))
        Hash(5) = SHA256_AddUnsigned(F, Hash(5))
        Hash(6) = SHA256_AddUnsigned(G, Hash(6))
        Hash(7) = SHA256_AddUnsigned(H, Hash(7))
    Next

    SHA256_Start = LCase$(Right$("00000000" & Hex(Hash(0)), 8) & _
                          Right$("00000000" & Hex(Hash(1)), 8) & _
                          Right$("00000000" & Hex(Hash(2)), 8) & _
                          Right$("00000000" & Hex(Hash(3)), 8) & _
                          Right$("00000000" & Hex(Hash(4)), 8) & _
                          Right$("00000000" & Hex(Hash(5)), 8) & _
                          Right$("00000000" & Hex(Hash(6)), 8) & _
                          Right$("00000000" & Hex(Hash(7)), 8))
'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'================================================================================='
Private Function SHA256_LShift( _
                 ByVal lValue As Long, _
                 ByVal iShiftBits As Integer _
        ) As Long
'---------------------------------------------------------------------------------'
    If iShiftBits = 0 Then
        SHA256_LShift = lValue
        Exit Function

    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            SHA256_LShift = &H80000000
        Else
            SHA256_LShift = 0
        End If
        Exit Function

    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And SHA256_m_l2Power(31 - iShiftBits)) Then
        SHA256_LShift = ((lValue And SHA256_m_lOnBits(31 - (iShiftBits + 1))) * _
            SHA256_m_l2Power(iShiftBits)) Or &H80000000
    Else
        SHA256_LShift = ((lValue And SHA256_m_lOnBits(31 - iShiftBits)) * _
            SHA256_m_l2Power(iShiftBits))
    End If
'---------------------------------------------------------------------------------'
End Function
'================================================================================='


'========================================================================='
Private Function SHA256_RShift( _
                 ByVal lValue As Long, _
                 ByVal iShiftBits As Integer _
        ) As Long
'-------------------------------------------------------------------------'
    If iShiftBits = 0 Then
        SHA256_RShift = lValue
        Exit Function

    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            SHA256_RShift = 1
        Else
            SHA256_RShift = 0
        End If
        Exit Function

    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    SHA256_RShift = (lValue And &H7FFFFFFE) \ SHA256_m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        SHA256_RShift = (SHA256_RShift Or (&H40000000 \ _
        SHA256_m_l2Power(iShiftBits - 1)))
    End If
'-------------------------------------------------------------------------'
End Function
'========================================================================='


'============================================================'
Private Function SHA256_AddUnsigned( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'------------------------------------------------------------'
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    SHA256_AddUnsigned = lResult
'------------------------------------------------------------'
End Function
'============================================================'


'=============================================='
Private Function SHA256_Ch( _
                 ByVal x As Long, _
                 ByVal Y As Long, _
                 ByVal z As Long _
        ) As Long
'----------------------------------------------'
    SHA256_Ch = ((x And Y) Xor ((Not x) And z))
'----------------------------------------------'
End Function
'=============================================='


'======================================================='
Private Function SHA256_Maj( _
                 ByVal x As Long, _
                 ByVal Y As Long, _
                 ByVal z As Long _
        ) As Long
'-------------------------------------------------------'
    SHA256_Maj = ((x And Y) Xor (x And z) Xor (Y And z))
'-------------------------------------------------------'
End Function
'======================================================='


'====================================================================================================================='
Private Function SHA256_s( _
                 ByVal x As Long, _
                 ByVal N As Long _
        ) As Long
'---------------------------------------------------------------------------------------------------------------------'
    SHA256_s = (SHA256_RShift(x, (N And SHA256_m_lOnBits(4))) Or SHA256_LShift(x, (32 - (N And SHA256_m_lOnBits(4)))))
'---------------------------------------------------------------------------------------------------------------------'
End Function
'====================================================================================================================='


'================================================================'
Private Function SHA256_R( _
                 ByVal x As Long, _
                 ByVal N As Long _
        ) As Long
'----------------------------------------------------------------'
    SHA256_R = SHA256_RShift(x, CInt(N And SHA256_m_lOnBits(4)))
'----------------------------------------------------------------'
End Function
'================================================================'


'==========================================================================='
Private Function SHA256_Sigma0( _
                 ByVal x As Long _
        ) As Long
'---------------------------------------------------------------------------'
    SHA256_Sigma0 = (SHA256_s(x, 2) Xor SHA256_s(x, 13) Xor SHA256_s(x, 22))
'---------------------------------------------------------------------------'
End Function
'==========================================================================='


'==========================================================================='
Private Function SHA256_Sigma1( _
                 ByVal x As Long _
        ) As Long
'---------------------------------------------------------------------------'
    SHA256_Sigma1 = (SHA256_s(x, 6) Xor SHA256_s(x, 11) Xor SHA256_s(x, 25))
'---------------------------------------------------------------------------'
End Function
'==========================================================================='


'==========================================================================='
Private Function SHA256_Gamma0( _
                 ByVal x As Long _
        ) As Long
'---------------------------------------------------------------------------'
    SHA256_Gamma0 = (SHA256_s(x, 7) Xor SHA256_s(x, 18) Xor SHA256_R(x, 3))
'---------------------------------------------------------------------------'
End Function
'==========================================================================='


'==========================================================================='
Private Function SHA256_Gamma1( _
                 ByVal x As Long _
        ) As Long
'---------------------------------------------------------------------------'
    SHA256_Gamma1 = (SHA256_s(x, 17) Xor SHA256_s(x, 19) Xor SHA256_R(x, 10))
'---------------------------------------------------------------------------'
End Function
'==========================================================================='


'======================================================================'
Private Function SHA256_ConvertToWordArray( _
                 ByRef sMessage As String _
        ) As Long()
'----------------------------------------------------------------------'
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long

    Const MODULUS_BITS    As Long = 512
    Const CONGRUENT_BITS  As Long = 448

    lMessageLength = Len(sMessage)

    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ SHA256_BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ SHA256_BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ SHA256_BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ SHA256_BYTES_TO_A_WORD

        lBytePosition = (3 - (lByteCount Mod SHA256_BYTES_TO_A_WORD)) _
                           * SHA256_BITS_TO_A_BYTE

        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))

        lWordArray(lWordCount) = lWordArray(lWordCount) Or _
                                 SHA256_LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ SHA256_BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod SHA256_BYTES_TO_A_WORD)) * _
                                         SHA256_BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or _
        SHA256_LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 1) = SHA256_LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = SHA256_RShift(lMessageLength, 29)

    SHA256_ConvertToWordArray = lWordArray
'----------------------------------------------------------------------'
End Function
'======================================================================'


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================='
Private Function SHA512_BSwap32( _
                 ByVal lX As Long _
        ) As Long
'-----------------------------------------------------------------------------'
    SHA512_BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 _
                  Or (lX And &HFF0000) \ &H100 Or _
                     (lX And &HFF000000) \ &H1000000 And &HFF Or _
                   -((lX And &H80) <> 0) * &H80000000
'-----------------------------------------------------------------------------'
End Function
'============================================================================='


'================================================================================'
Private Sub SHA512_pvAdd64( _
            ByRef lAL As Long, _
            ByRef lAH As Long, _
            ByVal lBL As Long, _
            ByVal lBH As Long _
        )
'--------------------------------------------------------------------------------'
    Dim lSign As Long

    If SHA512_m_bNoIntegerOverflowChecks Then
        lAL = lAL + lBL
        lAH = lAH + lBH
        If (lAL And &H80000000) <> 0 Then
            lSign = 1
        Else
            lSign = 0
        End If
        If (lBL And &H80000000) <> 0 Then
            lSign = lSign - 1
        End If
        Select Case True
        Case lSign < 0, lSign = 0 And (lAL And &H7FFFFFFF) < (lBL And &H7FFFFFFF)
            lAH = lAH + 1
        End Select
    Else
        If (lAL Xor lBL) >= 0 Then
            lAL = ((lAL Xor &H80000000) + lBL) Xor &H80000000
        Else
            lAL = lAL + lBL
        End If
        If (lAH Xor lBH) >= 0 Then
            lAH = ((lAH Xor &H80000000) + lBH) Xor &H80000000
        Else
            lAH = lAH + lBH
        End If
        If (lAL And &H80000000) <> 0 Then
            lSign = 1
        End If
        If (lBL And &H80000000) <> 0 Then
            lSign = lSign - 1
        End If
        Select Case True
        Case lSign < 0, lSign = 0 And (lAL And &H7FFFFFFF) < (lBL And &H7FFFFFFF)
            If lAH >= 0 Then
                lAH = ((lAH Xor &H80000000) + 1) Xor &H80000000
            Else
                lAH = lAH + 1
            End If
        End Select
    End If
'--------------------------------------------------------------------------------'
End Sub
'================================================================================'


'========================================================================================='
Private Function SHA512_pvSum0L( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'-----------------------------------------------------------------------------------------'
    SHA512_pvSum0L = ((lX And (SHA512_LNG_POW2_6 - 1)) * SHA512_LNG_POW2_25 _
        Or -((lX And SHA512_LNG_POW2_6) <> 0) * &H80000000) _
        Xor ((lX And (SHA512_LNG_POW2_1 - 1)) * SHA512_LNG_POW2_30 _
        Or -((lX And SHA512_LNG_POW2_1) <> 0) * &H80000000) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_28 Or -(lX < 0) * SHA512_LNG_POW2_3) _
        Xor ((lY And &H7FFFFFFF) \ SHA512_LNG_POW2_7 Or -(lY < 0) * SHA512_LNG_POW2_24) _
        Xor ((lY And &H7FFFFFFF) \ SHA512_LNG_POW2_2 Or -(lY < 0) * SHA512_LNG_POW2_29) _
        Xor ((lY And (SHA512_LNG_POW2_27 - 1)) * SHA512_LNG_POW2_4 _
        Or -((lY And SHA512_LNG_POW2_27) <> 0) * &H80000000)
'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'========================================================================================='
Private Function SHA512_pvSum1L( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'-----------------------------------------------------------------------------------------'
    SHA512_pvSum1L = ((lX And (SHA512_LNG_POW2_8 - 1)) * SHA512_LNG_POW2_23 _
        Or -((lX And SHA512_LNG_POW2_8) <> 0) * &H80000000) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_14 Or -(lX < 0) * SHA512_LNG_POW2_17) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_18 Or -(lX < 0) * SHA512_LNG_POW2_13) _
        Xor ((lY And &H7FFFFFFF) \ SHA512_LNG_POW2_9 Or -(lY < 0) * SHA512_LNG_POW2_22) _
        Xor ((lY And (SHA512_LNG_POW2_13 - 1)) * SHA512_LNG_POW2_18 _
        Or -((lY And SHA512_LNG_POW2_13) <> 0) * &H80000000) _
        Xor ((lY And (SHA512_LNG_POW2_17 - 1)) * SHA512_LNG_POW2_14 _
        Or -((lY And SHA512_LNG_POW2_17) <> 0) * &H80000000)
'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'================================================================================================='
Private Function SHA512_pvSig0L( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'-------------------------------------------------------------------------------------------------'
    SHA512_pvSig0L = ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_1 Or -(lX < 0) * SHA512_LNG_POW2_30) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_7 Or -(lX < 0) * SHA512_LNG_POW2_24) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_8 Or -(lX < 0) * SHA512_LNG_POW2_23) _
        Xor ((lY And 0) * SHA512_LNG_POW2_31 Or -((lY And 1) <> 0) * &H80000000) _
        Xor ((lY And (SHA512_LNG_POW2_6 - 1)) * SHA512_LNG_POW2_25 _
        Or -((lY And SHA512_LNG_POW2_6) <> 0) * &H80000000) _
        Xor ((lY And (SHA512_LNG_POW2_7 - 1)) * SHA512_LNG_POW2_24 _
        Or -((lY And SHA512_LNG_POW2_7) <> 0) * &H80000000)
'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'================================================================================================='
Private Function SHA512_pvSig0H( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'-------------------------------------------------------------------------------------------------'
    SHA512_pvSig0H = ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_1 Or -(lX < 0) * SHA512_LNG_POW2_30) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_7 Or -(lX < 0) * SHA512_LNG_POW2_24) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_8 Or -(lX < 0) * SHA512_LNG_POW2_23) _
        Xor ((lY And 0) * SHA512_LNG_POW2_31 Or -((lY And 1) <> 0) * &H80000000) _
        Xor ((lY And (SHA512_LNG_POW2_7 - 1)) * SHA512_LNG_POW2_24 _
        Or -((lY And SHA512_LNG_POW2_7) <> 0) * &H80000000)
'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'========================================================================================'
Private Function SHA512_pvSig1L( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'----------------------------------------------------------------------------------------'
    SHA512_pvSig1L = ((lX And (SHA512_LNG_POW2_28 - 1)) * SHA512_LNG_POW2_3 _
        Or -((lX And SHA512_LNG_POW2_28) <> 0) * &H80000000) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_6 Or -(lX < 0) * SHA512_LNG_POW2_25) _
        Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_19 Or -(lX < 0) * SHA512_LNG_POW2_12) _
        Xor ((lY And &H7FFFFFFF) \ SHA512_LNG_POW2_29 Or -(lY < 0) * SHA512_LNG_POW2_2) _
        Xor ((lY And (SHA512_LNG_POW2_5 - 1)) * SHA512_LNG_POW2_26 _
        Or -((lY And SHA512_LNG_POW2_5) <> 0) * &H80000000) _
        Xor ((lY And (SHA512_LNG_POW2_18 - 1)) * SHA512_LNG_POW2_13 _
        Or -((lY And SHA512_LNG_POW2_18) <> 0) * &H80000000)
'----------------------------------------------------------------------------------------'
End Function
'========================================================================================'


'======================================================================================================'
Private Function SHA512_pvSig1H( _
                 ByVal lX As Long, _
                 ByVal lY As Long _
        ) As Long
'------------------------------------------------------------------------------------------------------'
    SHA512_pvSig1H = ((lX And (SHA512_LNG_POW2_28 - 1)) * SHA512_LNG_POW2_3 _
                     Or -((lX And SHA512_LNG_POW2_28) <> 0) * &H80000000) _
                     Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_6 Or -(lX < 0) * SHA512_LNG_POW2_25) _
                     Xor ((lX And &H7FFFFFFF) \ SHA512_LNG_POW2_19 Or -(lX < 0) * SHA512_LNG_POW2_12) _
                     Xor ((lY And &H7FFFFFFF) \ SHA512_LNG_POW2_29 Or -(lY < 0) * SHA512_LNG_POW2_2) _
                     Xor ((lY And (SHA512_LNG_POW2_18 - 1)) * SHA512_LNG_POW2_13 _
                     Or -((lY And SHA512_LNG_POW2_18) <> 0) * &H80000000)
'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'========================================================================================================================'
Private Sub SHA512_pvRound( _
            ByVal lX00 As Long, ByVal lX01 As Long, ByVal lX02 As Long, ByVal lX03 As Long, _
            ByVal lX04 As Long, ByVal lX05 As Long, ByRef lX06 As Long, ByRef lX07 As Long, _
            ByVal lX08 As Long, ByVal lX09 As Long, ByVal lX10 As Long, ByVal lX11 As Long, _
            ByVal lX12 As Long, ByVal lX13 As Long, ByRef lX14 As Long, ByRef lX15 As Long, _
            ByRef uArray As SHA512_ArrayLong32, ByVal lIdx As Long, ByVal lJdx As Long _
        )
'------------------------------------------------------------------------------------------------------------------------'
    SHA512_pvAdd64 lX14, lX15, uArray.Item(lIdx), uArray.Item(lIdx + 1)
    SHA512_pvAdd64 lX14, lX15, SHA512_LNG_K(lJdx + lIdx), SHA512_LNG_K(lJdx + lIdx + 1)
    SHA512_pvAdd64 lX14, lX15, lX12 Xor (lX08 And (lX10 Xor lX12)), lX13 Xor (lX09 And (lX11 Xor lX13))
    SHA512_pvAdd64 lX14, lX15, SHA512_pvSum1L(lX08, lX09), SHA512_pvSum1L(lX09, lX08)
    SHA512_pvAdd64 lX06, lX07, lX14, lX15
    SHA512_pvAdd64 lX14, lX15, SHA512_pvSum0L(lX00, lX01), SHA512_pvSum0L(lX01, lX00)
    SHA512_pvAdd64 lX14, lX15, ((lX00 Or lX04) And lX02) Or (lX04 And lX00), ((lX01 Or lX05) And lX03) Or (lX05 And lX01)
'------------------------------------------------------------------------------------------------------------------------'
End Sub
'========================================================================================================================'


'========================================================================================'
Private Sub SHA512_pvStore( _
            ByRef uArray As SHA512_ArrayLong32, _
            ByVal lIdx As Long _
        )
'----------------------------------------------------------------------------------------'
    Dim lTL As Long
    Dim lTH As Long
    Dim lUL As Long
    Dim lUH As Long

    With uArray
        lTL = .Item(lIdx)
        lTH = .Item(lIdx + 1)
        SHA512_pvAdd64 lTL, lTH, .Item((lIdx + 18) And &H1F), .Item((lIdx + 19) And &H1F)
        lUL = SHA512_pvSig0L(.Item((lIdx + 2) And &H1F), .Item((lIdx + 3) And &H1F))
        lUH = SHA512_pvSig0H(.Item((lIdx + 3) And &H1F), .Item((lIdx + 2) And &H1F))
        SHA512_pvAdd64 lTL, lTH, lUL, lUH
        lUL = SHA512_pvSig1L(.Item((lIdx + 28) And &H1F), .Item((lIdx + 29) And &H1F))
        lUH = SHA512_pvSig1H(.Item((lIdx + 29) And &H1F), .Item((lIdx + 28) And &H1F))
        SHA512_pvAdd64 lTL, lTH, lUL, lUH
        .Item(lIdx) = lTL
        .Item(lIdx + 1) = lTH
    End With
'----------------------------------------------------------------------------------------'
End Sub
'========================================================================================'


'========================================================'
Private Function SHA512_pvGetOverflowIgnored() As Boolean
'--------------------------------------------------------'
    On Error GoTo EH
    If &H8000 - 1 <> 0 Then
        SHA512_pvGetOverflowIgnored = True
    End If
EH:
'--------------------------------------------------------'
End Function
'========================================================'


'==============================================================================================================================='
Private Sub SHA512_CryptoSha512Init( _
            ByRef uCtx As SHA512_CryptoSha512Context, _
            ByVal lBitSize As Long _
        )
'-------------------------------------------------------------------------------------------------------------------------------'
    Const FADF_AUTO As Long = 1
    Dim vElem       As Variant
    Dim lIdx        As Long
    Dim vSplit      As Variant
    Dim pDummy      As LongPtr

    If SHA512_LNG_K(0) = 0 Then
        For Each vElem In Split("D728AE22 428A2F98 23EF65CD 71374491 EC4D3B2F B5C0FBCF 8189DBBC E9B5DBA5 F348B538 3956C25B " & _
                                "B605D019 59F111F1 AF194F9B 923F82A4 DA6D8118 AB1C5ED5 A3030242 D807AA98 45706FBE 12835B01 " & _
                                "4EE4B28C 243185BE D5FFB4E2 550C7DC3 F27B896F 72BE5D74 3B1696B1 80DEB1FE 25C71235 9BDC06A7 " & _
                                "CF692694 C19BF174 9EF14AD2 E49B69C1 384F25E3 EFBE4786 8B8CD5B5 0FC19DC6 77AC9C65 240CA1CC " & _
                                "592B0275 2DE92C6F 6EA6E483 4A7484AA BD41FBD4 5CB0A9DC 831153B5 76F988DA EE66DFAB 983E5152 " & _
                                "2DB43210 A831C66D 98FB213F B00327C8 BEEF0EE4 BF597FC7 3DA88FC2 C6E00BF3 930AA725 D5A79147 " & _
                                "E003826F 06CA6351 0A0E6E70 14292967 46D22FFC 27B70A85 5C26C926 2E1B2138 5AC42AED 4D2C6DFC " & _
                                "9D95B3DF 53380D13 8BAF63DE 650A7354 3C77B2A8 766A0ABB 47EDAEE6 81C2C92E 1482353B 92722C85 " & _
                                "4CF10364 A2BFE8A1 BC423001 A81A664B D0F89791 C24B8B70 0654BE30 C76C51A3 D6EF5218 D192E819 " & _
                                "5565A910 D6990624 5771202A F40E3585 32BBD1B8 106AA070 B8D2D0C8 19A4C116 5141AB53 1E376C08 " & _
                                "DF8EEB99 2748774C E19B48A8 34B0BCB5 C5C95A63 391C0CB3 E3418ACB 4ED8AA4A 7763E373 5B9CCA4F " & _
                                "D6B2B8A3 682E6FF3 5DEFB2FC 748F82EE 43172F60 78A5636F A1F0AB72 84C87814 1A6439EC 8CC70208 " & _
                                "23631E28 90BEFFFA DE82BDE9 A4506CEB B2C67915 BEF9A3F7 E372532B C67178F2 EA26619C CA273ECE " & _
                                "21C0C207 D186B8C7 CDE0EB1E EADA7DD6 EE6ED178 F57D4F7F 72176FBA 06F067AA A2C898A6 0A637DC5 " & _
                                "BEF90DAE 113F9804 131C471B 1B710B35 23047D84 28DB77F5 40C72493 32CAAB7B 15C9BEBC 3C9EBE0A " & _
                                "9C100D4C 431D67C4 CB3E42B6 4CC5D4BE FC657E2A 597F299C 3AD6FAEC 5FCB6FAB 4A475817 6C44198C")
            SHA512_LNG_K(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        SHA512_m_bNoIntegerOverflowChecks = SHA512_pvGetOverflowIgnored
    End If
    With uCtx
        Select Case lBitSize Mod 1000
            Case 384
                vSplit = Split("C1059ED8 CBBB9D5D 367CD507 629A292A 3070DD17 9159015A F70E5939 152FECD8 FFC00B31 67332667 " & _
                               "68581511 8EB44A87 64F98FA7 DB0C2E0D BEFA4FA4 47B5481D")
            Case 512
                vSplit = Split("F3BCC908 6A09E667 84CAA73B BB67AE85 FE94F82B 3C6EF372 5F1D36F1 A54FF53A ADE682D1 510E527F " & _
                               "2B3E6C1F 9B05688C FB41BD6B 1F83D9AB 137E2179 5BE0CD19")
            Case Else
                Err.Raise vbObjectError, , "Invalid bit-size for SHA-512 (" & lBitSize & ")"
        End Select
        lIdx = 0
        For Each vElem In vSplit
            .State.Item(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        .NPartial = 0
        .NInput = 0
        .BitSize = lBitSize
        With .ArrayBytes
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 1
            .cLocks = 1
            .pvData = VarPtr(uCtx.Block.Item(0))
            .cElements = SHA512_LNG_BLOCKSZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.bytes), VarPtr(.ArrayBytes), LenB(pDummy))
    End With
'-------------------------------------------------------------------------------------------------------------------------------'
End Sub
'==============================================================================================================================='


'=========================================================================================================================================='
Private Sub SHA512_CryptoSha512Update( _
            ByRef uCtx As SHA512_CryptoSha512Context, _
            ByRef baInput() As Byte, _
            Optional ByVal Pos As Long, _
            Optional ByVal Size As Long = -1 _
        )
'------------------------------------------------------------------------------------------------------------------------------------------'
    Dim lAL  As Long
    Dim lAH  As Long
    Dim lBL  As Long
    Dim lBH  As Long
    Dim lCL  As Long
    Dim lCh  As Long
    Dim lDL  As Long
    Dim lDH  As Long
    Dim lEL  As Long
    Dim lEH  As Long
    Dim lFL  As Long
    Dim lFH  As Long
    Dim lGL  As Long
    Dim lGH  As Long
    Dim lHL  As Long
    Dim lHH  As Long
    Dim lIdx As Long
    Dim lJdx As Long

    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        .NInput = .NInput + Size
        If .NPartial > 0 And Size > 0 Then
            lIdx = SHA512_LNG_BLOCKSZ - .NPartial
            If lIdx > Size Then
                lIdx = Size
            End If
            Call CopyMemory(.bytes(.NPartial), baInput(Pos), lIdx)
            .NPartial = .NPartial + lIdx
            Pos = Pos + lIdx
            Size = Size - lIdx
        End If
        Do While Size > 0 Or .NPartial = SHA512_LNG_BLOCKSZ
            If .NPartial <> 0 Then
                .NPartial = 0
            ElseIf Size >= SHA512_LNG_BLOCKSZ Then
                Call CopyMemory(.bytes(0), baInput(Pos), SHA512_LNG_BLOCKSZ)
                Pos = Pos + SHA512_LNG_BLOCKSZ
                Size = Size - SHA512_LNG_BLOCKSZ
            Else
                Call CopyMemory(.bytes(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If

            For lIdx = 0 To UBound(.Block.Item) Step 2
                lAL = SHA512_BSwap32(.Block.Item(lIdx))
                .Block.Item(lIdx) = SHA512_BSwap32(.Block.Item(lIdx + 1))
                .Block.Item(lIdx + 1) = lAL
            Next
            lAL = .State.Item(0): lAH = .State.Item(1)
            lBL = .State.Item(2): lBH = .State.Item(3)
            lCL = .State.Item(4): lCh = .State.Item(5)
            lDL = .State.Item(6): lDH = .State.Item(7)
            lEL = .State.Item(8): lEH = .State.Item(9)
            lFL = .State.Item(10): lFH = .State.Item(11)
            lGL = .State.Item(12): lGH = .State.Item(13)
            lHL = .State.Item(14): lHH = .State.Item(15)
            lIdx = 0
            Do While lIdx < 2 * SHA512_LNG_ROUNDS
                lJdx = 0
                Do While lJdx < SHA512_LNG_BLOCKSZ \ 4
                    SHA512_pvRound lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, .Block, lJdx + 0, lIdx
                    SHA512_pvRound lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, .Block, lJdx + 2, lIdx
                    SHA512_pvRound lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, .Block, lJdx + 4, lIdx
                    SHA512_pvRound lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, .Block, lJdx + 6, lIdx
                    SHA512_pvRound lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, .Block, lJdx + 8, lIdx
                    SHA512_pvRound lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, .Block, lJdx + 10, lIdx
                    SHA512_pvRound lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, .Block, lJdx + 12, lIdx
                    SHA512_pvRound lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, .Block, lJdx + 14, lIdx
                    lJdx = lJdx + 16
                Loop
                lIdx = lIdx + 32
                If lIdx >= 2 * SHA512_LNG_ROUNDS Then
                    Exit Do
                End If
                For lJdx = 0 To 30 Step 2
                    SHA512_pvStore .Block, lJdx
                Next
            Loop
            SHA512_pvAdd64 .State.Item(0), .State.Item(1), lAL, lAH
            SHA512_pvAdd64 .State.Item(2), .State.Item(3), lBL, lBH
            SHA512_pvAdd64 .State.Item(4), .State.Item(5), lCL, lCh
            SHA512_pvAdd64 .State.Item(6), .State.Item(7), lDL, lDH
            SHA512_pvAdd64 .State.Item(8), .State.Item(9), lEL, lEH
            SHA512_pvAdd64 .State.Item(10), .State.Item(11), lFL, lFH
            SHA512_pvAdd64 .State.Item(12), .State.Item(13), lGL, lGH
            SHA512_pvAdd64 .State.Item(14), .State.Item(15), lHL, lHH
        Loop
    End With
'------------------------------------------------------------------------------------------------------------------------------------------'
End Sub
'=========================================================================================================================================='


'==================================================================='
Private Sub SHA512_CryptoSha512Finalize( _
            ByRef uCtx As SHA512_CryptoSha512Context, _
            ByRef baOutput() As Byte _
        )
'-------------------------------------------------------------------'
    Static B(0 To 1)    As Long
    Dim baPad()         As Byte
    Dim lIdx            As Long
    Dim pDummy          As LongPtr

    With uCtx
        lIdx = SHA512_LNG_BLOCKSZ - .NPartial
        If lIdx < 17 Then
            lIdx = lIdx + SHA512_LNG_BLOCKSZ
        End If
        ReDim baPad(0 To lIdx - 1) As Byte
        baPad(0) = &H80
        .NInput = .NInput / 10000@ * 8
        Call CopyMemory(B(0), .NInput, 8)
        Call CopyMemory(baPad(lIdx - 4), SHA512_BSwap32(B(0)), 4)
        Call CopyMemory(baPad(lIdx - 8), SHA512_BSwap32(B(1)), 4)
        SHA512_CryptoSha512Update uCtx, baPad
        Debug.Assert .NPartial = 0
        ReDim baOutput(0 To (.BitSize + 7) \ 8 - 1) As Byte
        .ArrayBytes.pvData = VarPtr(.State.Item(0))
        For lIdx = 0 To UBound(baOutput)
            baOutput(lIdx) = .bytes(lIdx + 7 - 2 * (lIdx And 7))
        Next
        Call CopyMemory(ByVal ArrPtr(.bytes), pDummy, LenB(pDummy))
    End With
'-------------------------------------------------------------------'
End Sub
'==================================================================='


'================================================================='
Private Function SHA512_CryptoSha512ByteArray( _
                 ByVal lBitSize As Long, _
                 ByRef baInput() As Byte, _
                 Optional ByVal Pos As Long, _
                 Optional ByVal Size As Long = -1 _
        ) As Byte()
'-----------------------------------------------------------------'
    Dim uCtx As SHA512_CryptoSha512Context
    SHA512_CryptoSha512Init uCtx, lBitSize
    SHA512_CryptoSha512Update uCtx, baInput, Pos, Size
    SHA512_CryptoSha512Finalize uCtx, SHA512_CryptoSha512ByteArray
'-----------------------------------------------------------------'
End Function
'================================================================='


'============================================================'
Private Function SHA512_ToHex( _
                 ByRef baData() As Byte _
        ) As String
'------------------------------------------------------------'
    Dim lIdx            As Long
    Dim sByte           As String

    SHA512_ToHex = String$(UBound(baData) * 2 + 2, 48)
    For lIdx = 0 To UBound(baData)
        sByte = LCase$(Hex$(baData(lIdx)))
        Mid$(SHA512_ToHex, lIdx * 2 + 3 - Len(sByte)) = sByte
    Next
'------------------------------------------------------------'
End Function
'============================================================'


'============================================================================================='
Private Function SHA512_CryptoSha512Text( _
                 ByVal lBitSize As Long, _
                 ByRef Data_Bytes() As Byte _
        ) As String
'---------------------------------------------------------------------------------------------'
    SHA512_CryptoSha512Text = SHA512_ToHex(SHA512_CryptoSha512ByteArray(lBitSize, Data_Bytes))
'---------------------------------------------------------------------------------------------'
End Function
'============================================================================================='


'========================================================'
Private Function SHA512_shl( _
                 ByVal Value As Long, _
                 ByVal Shift As Byte _
        ) As Long
'--------------------------------------------------------'
    SHA512_shl = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim M As Long
        For i = 1 To Shift
            M = SHA512_shl And &H40000000
            SHA512_shl = (SHA512_shl And &H3FFFFFFF) * 2
            If M <> 0 Then
                SHA512_shl = SHA512_shl Or &H80000000
            End If
        Next i
    End If
'--------------------------------------------------------'
End Function
'========================================================'


'-----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-----------------------------------------------------------------------------------------------------------------------------'


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

    '````````````````````````'
    Call Init_VBD_Kit_Hashing
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


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'

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

'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'
