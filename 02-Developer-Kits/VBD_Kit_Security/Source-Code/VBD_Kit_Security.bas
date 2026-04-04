Attribute VB_Name = "VBD_Kit_Security"
' | ================================================================================================== | '
' | ________              ______ ________               ________             ________           _____  | '
' | ___  __ \_____ __________  /___  ___/___________    ___  __ \__________________(_)____________  /_ | '
' | __  / / /  __ `/_  ___/_  //_/____ \_  _ \  ___/    __  /_/ /_  ___/  __ \____  /_  _ \  ___/  __/ | '
' | _  /_/ // /_/ /_  /   _  ,<  ____/ //  __/ /__      _  ____/_  /   / /_/ /___  / /  __/ /__ / /_   | '
' | /_____/ \__,_/ /_/    /_/|_| /____/ \___/\___/______/_/     /_/    \____/___  /  \___/\___/ \__/   | '
' |                                              _/_____/                    /___/                     | '
' | ================================================================================================== | '

' +-[MODULE: VBD_Kit_Security]------------------------------------------------------------------------------+
' |                                                                                                         |
' | [ENGINEER]: Zeus_0x01                                                                                   |
' | [TELEGRAM]: @Zeus_0x01 (Public Name)                                                                    |
' | [DESCRIPTION]: Реализация множества алгоритмов для управления настройками безопасности файлов MS Office |
' |                                                                                                         |
' +---------------------------------------------------------------------------------------------------------+

' // <copyright file="VBD_Kit_Security.bas" division="DarkSec_Project">
' // (C) Copyright 2023 Zeus_0x01 "{8C7D7042-0FD2-46B0-A55E-FEDCE95E762B}"
' // </copyright>

'-----------------------------------------------------------------------'
' // Implemented Functionality (Реализованный функционал):
'    Forum -  https://www.script-coding.ru/threads/vbd_kit_security.197
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
Private Const GUID_VBComponent As String = "{8C7D7042-0FD2-46B0-A55E-FEDCE95E762B}"
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

'------------------------------------------------------'
Public Enum Security_ManagementCenter_Privileges
    Without_PrivilegesChanges = &HFFFFFFFF
    Enable_AllMacros_NotRecommended = &H1
    Disable_AllMacros = &H4
    Disable_AllMacros_WithNotification = &H2
    Disable_AllMacros_ExceptDigitallySignedMacros = &H3
End Enum
'------------------------------------------------------'

'------------------------------------------------------'
Public Enum Security_ManagementCenter_AccessObjectModel
    Without_AOMChanges = &HFFFFFFFF
    Access_Denied = &H0
    Access_Provided = &H1
End Enum
'------------------------------------------------------'

'--------------------------------'
Public Enum Security_SettingMacro
    Privileges = &H0
    AccessObjectModel = &H1
End Enum
'--------------------------------'

'-------------------------------------'
Public Enum Type_Audit_ExcelObjectModel
    Audit_FullObjectModel = &H0
    Audit_Workbook = &H1
    Audit_Worksheet = &H2
End Enum
'-------------------------------------'

'-----------------------------------------'
Public Enum MSOffice_Type_Document_Group_1
    mso_Excel_1 = &H2
    mso_PowerPoint_1 = &H3
    mso_Word_1 = &H4
    mso_Access_1 = &H5
    mso_Outlook_1 = &H6
End Enum
'-----------------------------------------'

'-----------------------------------------'
Public Enum MSOffice_Type_Document_Group_2
    mso_Excel_2 = &H2
    mso_PowerPoint_2 = &H3
    mso_Word_2 = &H4
    mso_Access_2 = &H5
End Enum
'-----------------------------------------'

'-----------------------------------------'
Public Enum MSOffice_Type_Document_Group_3
    mso_Excel_3 = &H2
    mso_PowerPoint_3 = &H3
    mso_Word_3 = &H4
End Enum
'-----------------------------------------'

'--------------------------------------'
Public Enum MSOffice_Type_ProtectedView
    mso_AttachmentsInPV = &H0
    mso_InternetFilesInPV = &H1
    mso_UnsafeLocationsInPV = &H2
End Enum
'--------------------------------------'

'---------------------------------------'
Public Enum MSExcel_Type_ExternalContent
    mse_DataConnectionWarnings = &H0
    mse_WorkbookLinkWarnings = &H1
End Enum
'---------------------------------------'

'-------------------------------------------------------'
Public Enum MSExcel_Setting_ExternalContent
    Enable_AllUpdatesAndConnections_NotRecommended = &H0
    Request_ForUpdatesAndConnections = &H1
    Disable_AllUpdatesAndConnections = &H2
End Enum
'-------------------------------------------------------'

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

'------------------------------------'
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
'------------------------------------'

'---------------------------------------------------'
Private Enum WinAPI_VirtualProtect_MemoryProtection
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_TARGETS_NO_UPDATE = &H40000000
End Enum
'---------------------------------------------------'

'-----------------------------------------'
Private Enum MSOffice_Type_ArrayDimensions
    Array_Error = &HFFFFFFFF
    Array_NotInitialized = &H0
    Array_OneDimensional = &H1
    Array_TwoDimensional = &H2
    Array_MultiDimensional = &H3
End Enum
'-----------------------------------------'

'----------------------------------------------'
Private Enum MSOffice_Type_ContentFormat_Sec
    MSDI_Undefined_Sec = &HFFFFFFFC
    MSDI_Null_Sec = &HFFFFFFFD
    MSDI_NullString_Sec = &HFFFFFFFE
    MSDI_Nothing_Sec = &HFFFFFFFF
    MSDI_Empty_Sec = &H0
    MSDI_Bool_Sec = &H1
    MSDI_Int8_Sec = &H2
    MSDI_Int16_Sec = &H4
    MSDI_Int32_Sec = &H8
    MSDI_Int64_Sec = &H10
    MSDI_FloatPoint32_Sec = &H20
    MSDI_FloatPoint64Cur_Sec = &H40
    MSDI_FloatPoint64Dbl_Sec = &H56
    MSDI_FloatPoint112_Sec = &H80
    MSDI_Date_Sec = &H100
    MSDI_String_Sec = &H200
    MSDI_Range_Sec = &H400
    MSDI_Array_Sec = &H800
    MSDI_Object_Sec = &H1000
    MSDI_Collection_Sec = &H2000
    MSDI_Dictionary_Sec = &H4000
    MSDI_File_Sec = &H8000
    MSDI_Folder_Sec = &H10000
    MSDI_VBComponent_Sec = &H20000
    MSDI_VBProject_Sec = &H40000
    MSDI_UserDefType_Sec = &H80000
    MSDI_NonExistent_Directory_Sec = &H100000
    
    ' {
        MSDI_Procedure_Sec = &H120000
        MSDI_AutoDetect_Sec = &HFFFFFFAD
    ' }
End Enum
'----------------------------------------------'

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

'--------------------------------------'
Private Enum FileSystem_Directory_Type
    DirType_File = &H0
    DirType_Folder = &H1
    DirType_Invalid = &HFFFFFFFF
    DirType_NotFound = &HFFFFFFFE
End Enum
'--------------------------------------'

'----------------------------------------------'
Private Enum FileSystem_FileDialog_Filters_TID
    TID_FileSystem_AllFiles = &H0
    TID_FileSystem_BIN = &H1
    TID_FileSystem_DLL = &H2
    TID_FileSystem_EXE = &H3
    TID_FileSystem_INI = &H4
    TID_FileSystem_TXT = &H5
    TID_FileSystem_Access = &H6
    TID_FileSystem_Excel = &H7
    TID_FileSystem_PowerPoint = &H8
    TID_FileSystem_Word = &H9
End Enum
'----------------------------------------------'

'------------------------------'
Private Enum UI_Controls_Type
    Control_CheckBox = &H0
    Control_ComboBox = &H1
    Control_CommandButton = &H2
    Control_Frame = &H3
    Control_Image = &H4
    Control_Label = &H5
    Control_ListBox = &H6
    Control_MultiPage = &H7
    Control_OptionButton = &H8
    Control_RefEdit = &H9
    Control_ScrollBar = &HA
    Control_SpinButton = &HB
    Control_TabStrip = &HC
    Control_TextBox = &HD
    Control_ToggleButton = &HE
End Enum
'------------------------------'

'-------------------------------------'
Private Type FileSystem_MetaData_File
    MD_File_Name            As String
    MD_File_TempName        As String
    MD_File_BaseName        As String
    MD_File_TypeName        As String
    MD_Folder_TempName      As String
    MD_Folder_TempDirectory As String
End Type
'-------------------------------------'

'------------------------'
Private Type WinAPI_GUID
    Data_1 As Long
    Data_2 As Integer
    Data_3 As Integer
    Data_4(7) As Byte
End Type
'------------------------'

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

'-----------------------------------'
Private Type Registry_KeyData
    Key_Name  As String
    Key_Type  As Registry_ValueTypes
    Key_Value As Variant
End Type
'-----------------------------------'

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

'---------------------------------'
Private Type UI_UserForm_Property
    UI_Name        As String
    UI_Caption     As String
    UI_Width       As Double
    UI_Height      As Double
    UI_ShowModal   As Boolean
    UI_ShowBorders As Boolean
    UI_BackColor   As Variant
    UI_BorderColor As Variant
    UI_CodeModule  As String
End Type
'---------------------------------'

'--------------------------------------------'
Private Type UI_Control_Property
    UI_Control_Type      As UI_Controls_Type
    UI_Control_Name      As String
    UI_Control_Caption   As String
    UI_Control_Width     As Double
    UI_Control_Height    As Double
    UI_Control_Top       As Double
    UI_Control_Left      As Double
    UI_Control_FontSize  As Double
    UI_Control_FontBold  As Boolean
    UI_Control_BackColor As Variant
    UI_Control_ForeColor As Variant
End Type
'--------------------------------------------'

'-----------------------------------'
Private Type Change_ContextVBProject
    Position As Long
    Value    As Long
End Type
'-----------------------------------'

'--------------------------------'
#Const Has_PtrSafe = (VBA7 <> 0&)
'--------------------------------'

'------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (Kernel32.dll)

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
            Function VirtualProtect Lib "kernel32.dll" ( _
                     ByRef lpAddress As LongPtr, _
                     ByVal dwSize As LongPtr, _
                     ByVal flNewProtect As LongPtr, _
                     lpflOldProtect As LongPtr _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetModuleHandleA Lib "kernel32.dll" ( _
                     ByVal lpModuleName As String _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetProcAddress Lib "kernel32.dll" ( _
                     ByVal hModule As LongPtr, _
                     ByVal lpProcName As String _
            ) As LongPtr

    Private Declare PtrSafe _
            Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                ByRef Destination As LongPtr, _
                ByRef Source As LongPtr, _
                ByVal Length As LongPtr _
            )

    Private Declare PtrSafe _
            Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                ByRef Destination As Any, _
                ByRef Source As Any, _
                ByVal Length As Long _
            )

    Private Declare PtrSafe _
            Function VirtualAlloc Lib "kernel32.dll" ( _
                     ByVal lpAddress As LongPtr, _
                     ByVal dwSize As LongPtr, _
                     ByVal flAllocationType As Long, _
                     ByVal flProtect As Long _
            ) As LongPtr

    Private Declare PtrSafe _
            Function VirtualFree Lib "kernel32.dll" ( _
                     ByVal lpAddress As LongPtr, _
                     ByVal dwSize As LongPtr, _
                     ByVal dwFreeType As Long _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then  ' // Windows API (Advapi32.dll)

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
            Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal dwIndex As Long, _
                     ByVal lpName As String, _
                     ByRef lpcbName As Long, _
                     ByVal lpReserved As LongPtr, _
                     ByVal lpClass As LongPtr, _
                     ByRef lpClass As LongPtr, _
                     ByRef pftLastWriteTime As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal dwIndex As Long, _
                     ByVal lpValueName As String, _
                     ByRef lpcbValueName As Long, _
                     ByVal lpReserved As LongPtr, _
                     ByVal lpType As LongPtr, _
                     ByVal lpData As Byte, _
                     ByVal lpcbData As LongPtr _
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
            Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpSubKey As String, _
                     ByVal Reserved As Long, _
                     ByVal lpClass As String, _
                     ByVal dwOptions As Long, _
                     ByVal samDesired As Long, _
                     ByVal lpSecurityAttributes As LongPtr, _
                     ByRef phkResult As LongPtr, _
                     ByRef lpdwDisposition As Long _
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
            Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpSubKey As String _
            ) As Long

    Private Declare PtrSafe _
            Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
                     ByVal HKey As LongPtr, _
                     ByVal lpValueName As String _
            ) As Long

    Private Declare PtrSafe _
            Function RegCloseKey Lib "advapi32.dll" ( _
                     ByVal HKey As LongPtr _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (User32.dll)

    Private Declare PtrSafe _
            Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                     ByVal lpClassName As String, _
                     ByVal lpWindowName As String _
            ) As LongPtr

    Private Declare PtrSafe _
            Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
                     ByVal hWnd1 As LongPtr, _
                     ByVal hwnd2 As LongPtr, _
                     ByVal lpsz1 As String, _
                     ByVal lpsz2 As String _
            ) As LongPtr

    Private Declare PtrSafe _
            Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                     ByVal hWnd As LongPtr, _
                     ByVal wMsg As Long, _
                     ByVal wParam As LongPtr, _
                     ByVal lParam As String _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetNextWindow Lib "user32.dll" Alias "GetWindow" ( _
                     ByVal hWnd As LongPtr, _
                     ByVal wFlag As Long _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetDlgItem Lib "user32.dll" ( _
                     ByVal hDlg As LongPtr, _
                     ByVal nIDDlgItem As Long _
            ) As LongPtr

    Private Declare PtrSafe _
            Function DialogBoxParam Lib "user32.dll" Alias "DialogBoxParamA" ( _
                     ByVal hInstance As LongPtr, _
                     ByVal pTemplateName As LongPtr, _
                     ByVal hWndParent As LongPtr, _
                     ByVal lpDialogFunc As LongPtr, _
                     ByVal dwInitParam As LongPtr _
            ) As Integer

    Private Declare PtrSafe _
            Function CallDlgBxParam Lib "user32.dll" Alias "CallWindowProcW" ( _
                     ByVal pFunc As LongPtr, _
                     ByVal hInstance As LongPtr, _
                     ByVal pTemplateName As LongPtr, _
                     ByVal hWndParent As LongPtr, _
                     ByVal lpDialogFunc As LongPtr, _
                     ByVal dwInitParam As LongPtr _
            ) As LongPtr

#Else

    ' // Old MS Office or MacOS

#End If
'------------------------------------------------------------------------------------'

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

'---------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then ' // Windows API (OLE32.dll)

    Private Declare PtrSafe _
            Function CoCreateGuid Lib "OLE32.dll" ( _
                     GUID As WinAPI_GUID _
            ) As LongPtr

    Private Declare PtrSafe _
            Function StringFromGUID2 Lib "OLE32.dll" ( _
                     GUID As WinAPI_GUID, _
                     ByVal lpStrGuid As LongPtr, _
                     ByVal cbMax As Long _
            ) As LongPtr

#Else

    ' // Old MS Office or MacOS

#End If
'---------------------------------------------------------------'

'-----------------------------'
Private Glb_UI_Inx_Max As Long
'----------------------------------------------------------------------'
Private Gbl_ProgressBar_Width As Double, Gbl_ProgressBar_Text As String
'----------------------------------------------------------------------'

'-----------------------------------------'
Private Glb_Ptr_Func          As LongPtr
Private Glb_DefaultAddress    As LongPtr
Private Glb_TrampolineAddress As LongPtr
'-----------------------------------------'
Private Glb_Hooking_Bytes(0 To 11) As Byte
Private Glb_Default_Bytes(0 To 11) As Byte
'-----------------------------------------'

'------------------------------------------------------------------'
Private Glb_MSOffice_Type_Building    As MSOffice_Type_Building
Private Glb_MSOffice_Type_Application As MSOffice_Type_Application
'------------------------------------------------------------------'

'---------------------------------------------------------------------'
Private Const CDP_HOOK_ACTIVE     As String = "Hook_ActiveFlag"
Private Const CDP_TRAMPOLINE_ADDR As String = "Hook_TrampolineAddress"
'---------------------------------------------------------------------'
Private Const CDP_DEFAULT_ADDR    As String = "UnHook_DefaultAddress"
Private Const CDP_DEFAULT_BYTES   As String = "UnHook_DefaultBytes"
'---------------------------------------------------------------------'

'-------------------------------------------------------------------------------'
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000

Private Const LANG_NEUTRAL    As Long = &H0
Private Const SUBLANG_DEFAULT As Long = &H1
Private Const LANG_DEFAULT    As Long = (SUBLANG_DEFAULT * &H400 + LANG_NEUTRAL)
'-------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
Private Const REGISTRY_SECTION_MSO As String = "SOFTWARE\Microsoft\Office\"

Private Const REGISTRY_SECTION_EXCEL_SECURITY      As String = "\Excel\Security\"
Private Const REGISTRY_SECTION_WORD_SECURITY       As String = "\Word\Security\"
Private Const REGISTRY_SECTION_ACCESS_SECURITY     As String = "\Access\Security\"
Private Const REGISTRY_SECTION_OUTLOOK_SECURITY    As String = "\Outlook\Security\"
Private Const REGISTRY_SECTION_POWERPOINT_SECURITY As String = "\PowerPoint\Security\"
'-------------------------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------------------'
Private Const FILE_SYSTEM_LOCAL_APP_DATA                As String = "\AppData\Local"

Private Const FILE_SYSTEM_SAVE_PATH                     As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\"

Private Const FILE_SYSTEM_SAVE_PATH_UNPROTECT_BOOK      As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Unprotect_Books\"
Private Const FILE_SYSTEM_SAVE_PATH_PROTECTED_BOOK      As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Protected_Books\"

Private Const FILE_SYSTEM_SAVE_PATH_ENCRYPTED_DOCUMENT  As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Encrypted_Documents\"
Private Const FILE_SYSTEM_SAVE_PATH_DECRYPTED_DOCUMENT  As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Decrypted_Documents\"

Private Const FILE_SYSTEM_SAVE_PATH_UNPROTECT_VBPROJECT As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Unprotect_VBProject\"
Private Const FILE_SYSTEM_SAVE_PATH_PROTECTED_VBPROJECT As String = _
              "\DarkSec_Project\[VBD_Kit] Компоненты разработчика\VBD_Kit_Security\Protected_VBProject\"
'--------------------------------------------------------------------------------------------------------'


'================================================================================================================='
Public Function Audit_ExcelObjectModel( _
                Optional ByRef HackType_Protection_Excel As Type_Audit_ExcelObjectModel = Audit_FullObjectModel, _
                Optional ByRef FileSystem_FilePath As Variant, _
                Optional ByVal Show_List As Boolean = False _
       ) As Variant
'-----------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-----------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````'
    Dim File_Type        As MSOffice_Type_ContentFormat_Sec
    Dim File_Name        As String, File_FullName   As String
    Dim File_TempName    As String, File_TypeName   As String
    Dim File_BaseName    As String, File_PathFolder As String
    Dim tmp_Name         As String, tmp_FullName    As String
    Dim tmp_FolderName   As String, Error_Message   As String
    Dim ZIP_Archive_File As String, ZIP_FolderName  As String
    Dim Path_BookXML     As String, File_SheetXML   As String
    Dim Path_SheetXML    As String, Coll_FilesName  As New Collection
    '````````````````````````````````````````````````````````````````````````````````'
    Dim FileSystem_Type_Data As FileSystem_Directory_Type
    Dim Handle_FreeFile      As Integer, xmlFileContent As Variant
    Dim Processed_File       As Variant, UI_UserForm    As Object
    '````````````````````````````````````````````````````````````````````````````````'
    Dim Obj_FSO As Object, Obj_WB         As Object
    Dim Obj_SDI As Object, Obj_WS         As Object
    Dim Exl_App As Object, Native_Process As Boolean
    '````````````````````````````````````````````````````````````````````````````````'
    Dim Table_Name   As String, StatusBar As String, Buffer As String, File As Variant
    '````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_Files As Object, Obj_Folder     As Object, D_Key     As Variant
    Dim tmp_Arr    As Variant, Table_Headers As Variant, Inx_File As Long
    '````````````````````````````````````````````````````````````````````````````````'
    Dim Table_ColumnWidth As Variant, Count_Files As Long, Matrix_Result As Variant
    '````````````````````````````````````````````````````````````````````````````````'
    Dim Inx_LB1 As Long, Inx_UB1 As Long
    Dim i       As Long, Inx     As Long
    '````````````````````````````````````````````````````````````````````````````````'
    Const xl_OpenXMLWorkbookMacroEnabled As Long = 52&, xl_Excel12 As Long = 50&
    '````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````'
    Audit_ExcelObjectModel = False
    '````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Glb_MSOffice_Type_Application
        Case Type_MSOffice_Excel, Type_MSOffice_PowerPoint, Type_MSOffice_Word: Native_Process = True
        Case Else: Native_Process = False
    End Select

    If IsMissing(FileSystem_FilePath) Then
        FileSystem_FilePath = Get_List_Files_ToProcess(True, TID_FileSystem_Excel)
        If FileSystem_FilePath(1) = Empty Then Exit Function Else Inx = 1: GoTo GT¦Data_Handling
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    File_Type = Get_SemanticDataType(FileSystem_FilePath)
    Call Normalize_DataByType(FileSystem_FilePath, File_Type, Inx)
    '`````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Data_Handling:

    If Inx = 0 Then Audit_ExcelObjectModel = False: Exit Function

    Inx_LB1 = LBound(FileSystem_FilePath, 1)
    Inx_UB1 = UBound(FileSystem_FilePath, 1)

    Set Obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Dict_Files = CreateObject("Scripting.Dictionary")

    If Native_Process Then Let Glb_UI_Inx_Max = Inx_UB1: Set UI_UserForm = Create_UI_tmpForm()

    For i = Inx_LB1 To Inx_UB1

        File_FullName = FileSystem_FilePath(i)

        If Len(File_FullName) = 0 Then File_FullName = "-"

        If Dict_Files.Exists(File_FullName) Then GoTo GT¦Next_Loop
        FileSystem_Type_Data = Get_Directory_Type(File_FullName)

        Select Case FileSystem_Type_Data

            Case DirType_File

                If File_IsOpen(File_FullName) Then
                    Dict_Files(File_FullName) = "The File was Open!": GoTo GT¦Next_Loop
                Else
                    Processed_File = File_FullName
                    If Not Right(Processed_File, 4) = ".xls" Then
                        Dict_Files(File_FullName) = "Ready": GoSub GS¦Hack_Protection_File
                    Else
                        Dict_Files(File_FullName) = "Not supported"
                    End If
                End If

            Case DirType_Folder

                Set Obj_Folder = Obj_FSO.GetFolder(File_FullName)
                Count_Files = Obj_Folder.Files.Count: Inx_File = 0

                For Each File In Obj_Folder.Files

                    File_FullName = File.Path: Inx_File = Inx_File + 1

                    If Not Dict_Files.Exists(File_FullName) Then
                        If Not Obj_FSO.GetExtensionName(File_FullName) Like "xls*" Then
                            Dict_Files(File_FullName) = "The File not Support!"
                        Else
                            If File_IsOpen(File_FullName) Then
                                Dict_Files(File_FullName) = "The File was Open!"
                            Else
                                Processed_File = File_FullName
                                If Not Right(Processed_File, 4) = ".xls" Then
                                    Dict_Files(File_FullName) = "Ready": GoSub GS¦Hack_Protection_File
                                Else
                                    Dict_Files(File_FullName) = "Not supported"
                                End If
                            End If
                        End If
                    End If

                    If Native_Process Then
                        If Not Update_UI_tmpForm_ProgressBar_Lower(Inx_File, Count_Files, UI_UserForm.Name) Then
                            Call Remove_UI_tmpForm(UI_UserForm): GoTo GT¦Terminate_Function
                        End If
                        DoEvents
                    End If

                Next File

            Case DirType_NotFound

                Dict_Files(File_FullName) = "Directory or File Not Found!"

            Case DirType_Invalid

                Dict_Files(File_FullName) = "String is not defined as a Directory!"

        End Select

GT¦Next_Loop:

        If Native_Process Then
            If Not Update_UI_tmpForm_ProgressBar_Upper(i, UI_UserForm.Name) Then
                Call Remove_UI_tmpForm(UI_UserForm): GoTo GT¦Terminate_Function
            End If
            DoEvents
        End If

    Next i

    If Native_Process Then Call Remove_UI_tmpForm(UI_UserForm): Set UI_UserForm = Nothing
'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Terminate_Function:

    If Dict_Files.Count = 0 Then Audit_ExcelObjectModel = False: GoTo GT¦Clearing_Memory

    ReDim Matrix_Result(1 To Dict_Files.Count, 1 To 4): Inx = 0

    For Each D_Key In Dict_Files.keys
        Inx = Inx + 1: tmp_Arr = Split(Dict_Files(D_Key), "#$#")
        Matrix_Result(Inx, 1) = Inx
        Matrix_Result(Inx, 2) = D_Key
        Matrix_Result(Inx, 3) = tmp_Arr(0)
        If UBound(tmp_Arr, 1) > 0 Then Matrix_Result(Inx, 4) = tmp_Arr(1) Else Matrix_Result(Inx, 4) = "-"
    Next D_Key

    If Show_List Then

        Table_Name = "Unprotect_Books"
        Table_Headers = Array("INDEX" & vbCrLf & "Unique ID", _
                              "FILE NAME" & vbCrLf & "Documents FullName", _
                              "FILE STATUS" & vbCrLf & "File Final Status", _
                              "FILE LOCATION" & vbCrLf & "Path to folder file (Document)")
        Table_ColumnWidth = Array(12, 45, 32, 100)

        If Application.Name = "Microsoft Excel" Then
            If Not Create_WorkSheet_InExcel( _
                   Matrix_Result, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_UnprotectBooks", "objSheet_UnprotectBooks" _
            ) Then
                Error_Message = "Книга, в которую вы хотите вывести лог, защищена!" & _
                                " Создание новых листов возможно только после снятия защиты!"

                Show_ErrorMessage_Immediate Error_Message, "Структура книги защищена!"
            End If
        Else
            If Not Create_Process_ExcelApplication( _
                   Matrix_Result, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_UnprotectBooks" _
            ) Then
                Error_Message = "Невозможно создать отдельный процесс приложения посредством SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Непредвиденная ошибка интерфейса SDI!"
            End If
        End If

    End If

    If Not Coll_FilesName.Count = 0 Then FileSystem_Select_FilesInExplorer File_PathFolder, Coll_FilesName

    Audit_ExcelObjectModel = Matrix_Result
'`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````'
GT¦Clearing_Memory:

    If Not Exl_App Is Nothing Then Exl_App.Quit: Set Exl_App = Nothing
    Dir Environ$("TEMP")

    Exit Function
'`````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Hack_Protection_File:

    With Obj_FSO
        File_Name = .GetFileName(File_FullName)
        File_TempName = Left$(.GetTempName(), 8)
        File_BaseName = .GetBaseName(File_FullName)
        File_TypeName = .GetExtensionName(File_FullName)
    End With

    tmp_FullName = Environ$("Temp") & "\" & File_BaseName & "_" & _
                                            File_TempName & "_" & ".zip"

    tmp_FolderName = File_TempName & "_" & _
                     Format(Now, "DD.MM.YYYY") & " [" & _
                     Format(Now, "HH-MM-SS") & "]"

    tmp_FullName = Environ$("Temp") & "\" & File_BaseName & "_" & _
                                            File_TempName & "_" & ".zip"

    tmp_FolderName = File_TempName & "_" & _
                     Format(Now, "DD.MM.YYYY") & " [" & _
                     Format(Now, "HH-MM-SS") & "]"

    ZIP_FolderName = Environ$("TEMP") & "\" & tmp_FolderName
    File_PathFolder = Environ$("USERPROFILE") & FILE_SYSTEM_LOCAL_APP_DATA & _
                                                FILE_SYSTEM_SAVE_PATH_UNPROTECT_BOOK

    If Dir(ZIP_FolderName, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, ZIP_FolderName, ByVal 0&)
    End If

    If Dir(File_PathFolder, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, File_PathFolder, ByVal 0&)
    End If

    If File_TypeName = "xlsb" Then GoSub GS¦ReSave_Excel

    Obj_FSO.CopyFile File_FullName, tmp_FullName
    Call ZIP_Archive_UnPack(ZIP_FolderName, tmp_FullName)

    Select Case HackType_Protection_Excel
        Case 0: GoSub GS¦Audit_Workbook
                GoSub GS¦Audit_Worksheet
        Case 1: GoSub GS¦Audit_Workbook
        Case 2: GoSub GS¦Audit_Worksheet
    End Select

    ZIP_Archive_File = File_PathFolder & File_BaseName & ".zip"

    Call ZIP_Archive_Create(CVar(ZIP_Archive_File))
    Call ZIP_Archive_UnPack(ZIP_Archive_File, ZIP_FolderName & "\")

    With Obj_FSO

        .DeleteFolder ZIP_FolderName: .DeleteFile tmp_FullName
        If .FileExists(File_PathFolder & File_BaseName & "." & File_TypeName) Then
            .DeleteFile File_PathFolder & File_BaseName & "." & File_TypeName
        End If

        On Error Resume Next
            Do
                DoEvents
                .GetFile(ZIP_Archive_File).Name = File_BaseName & "." _
                                                                & File_TypeName
            Loop Until CBool(Err)
        On Error GoTo 0

        If Right(Processed_File, 4) = "xlsb" Then
            .DeleteFile tmp_Name & ".xlsm"
            File_FullName = File_PathFolder & File_BaseName & "." & File_TypeName
            GoSub GS¦ReSave_Excel
        End If

    End With

    On Error Resume Next
        Coll_FilesName.Add File_BaseName & "." & File_TypeName, File_PathFolder _
                                         & File_BaseName & "." & File_TypeName
    On Error GoTo 0

    Dict_Files(Processed_File) = "Ready#$#" & File_PathFolder _
                                            & File_BaseName & "." & File_TypeName

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦Audit_Workbook:

    Path_BookXML = ZIP_FolderName & "\xl\workbook.xml"

    Handle_FreeFile = FreeFile: Open Path_BookXML For Input As Handle_FreeFile
    xmlFileContent = Input(LOF(Handle_FreeFile), Handle_FreeFile): Close Handle_FreeFile

    Buffer = xmlFileContent

    Call Remove_ItemXML(Buffer, "workbookProtection")
    Call Remove_ItemXML(Buffer, "fileSharing")

    If Not StrComp(Buffer, xmlFileContent) = 0 Then
        Handle_FreeFile = FreeFile
        Open Path_BookXML For Output As #Handle_FreeFile
        Print #Handle_FreeFile, Buffer
        Close #Handle_FreeFile
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````'
GS¦Audit_Worksheet:

    File_SheetXML = Dir(ZIP_FolderName & "\xl\worksheets\*.xml*")

    Do While File_SheetXML <> ""

        Path_SheetXML = ZIP_FolderName & "\xl\worksheets\" & File_SheetXML
        Handle_FreeFile = FreeFile

        Open Path_SheetXML For Input As Handle_FreeFile
        xmlFileContent = Input(LOF(Handle_FreeFile), Handle_FreeFile)
        Close Handle_FreeFile

        Buffer = xmlFileContent: Call Remove_ItemXML(Buffer, "sheetProtection")

        If Not StrComp(Buffer, xmlFileContent) = 0 Then
            Handle_FreeFile = FreeFile
            Open Path_SheetXML For Output As #Handle_FreeFile
            Print #Handle_FreeFile, Buffer
            Close #Handle_FreeFile
        End If

        File_SheetXML = Dir

    Loop

    Return
'``````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````'
GS¦ReSave_Excel:

    If Exl_App Is Nothing Then
        Set Exl_App = CreateObject("Excel.Application")
        With Exl_App
            .DisplayAlerts = False: .ScreenUpdating = False
        End With
    End If

    Set Obj_WB = Exl_App.Workbooks.Open(File_FullName, False, True)
    tmp_Name = File_PathFolder & File_BaseName & "_" & File_TempName

    If File_TypeName = "xlsb" Then
        Obj_WB.SaveAs tmp_Name & ".xlsm", _
        FileFormat:=xl_OpenXMLWorkbookMacroEnabled
        File_FullName = tmp_Name & ".xlsm": File_TypeName = "xlsm"
    Else
        Obj_WB.SaveAs File_PathFolder & File_BaseName & ".xlsb", _
        FileFormat:=xl_Excel12: File_TypeName = "xlsb"
        Kill File_FullName
    End If

    Obj_WB.Close SaveChanges:=False

    Return
'`````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================'


'================================================================================================================'
Public Function Audit_OfficeCrypto() As Variant
'----------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox _
    "Данная функция (""Audit_OfficeCrypto"") не включена в текущую сборку модуля по инциативе разработчика!", _
                                                                          vbInformation, "[DarkSec_Project]"
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================'


'========================================================================================================================'
Public Function Audit_Unviewable( _
                Optional ByVal FileSystem_FilePath As Variant _
       ) As Boolean
'------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'------------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````'
    Dim File_MetaData          As FileSystem_MetaData_File
    Dim File_PathFolder        As String, File_VBProject As String
    Dim FileSystem_File_Format As String, File_Name      As String
    '`````````````````````````````````````````````````````````````'
    Dim Buffer  As String, FF       As Integer
    Dim Obj_FSO As Object, Obj_File As Object
    '`````````````````````````````````````````````````````````````'
    Dim ZIP_Archive_File  As String, Obj_App        As Object
    Dim Source_VBAProject As String, Obj_FileOffice As Object
    Dim Coll_FilesName    As New Collection
    Dim tmp_FilePath      As String, Flag_Result    As Boolean
    '`````````````````````````````````````````````````````````````'
    Dim tmp_Arr As Variant, Error_Message As String
    '`````````````````````````````````````````````````````````````'

    '```````````````````````'
    Audit_Unviewable = False
    '```````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    If IsMissing(FileSystem_FilePath) Then
        FileSystem_FilePath = Get_List_Files_ToProcess(False, TID_FileSystem_Excel, _
                                                              TID_FileSystem_Word, _
                                                              TID_FileSystem_PowerPoint)
        If Len(FileSystem_FilePath(1)) = 0 Then Exit Function
    Else
        If IsArray(FileSystem_FilePath) Or IsObject(FileSystem_FilePath) Then
            Error_Message = "Поддержка массивов и объектов не реализована!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории!"
            Exit Function
        End If

        If Not Get_Directory_Type(CStr(FileSystem_FilePath)) = DirType_File Then
            Error_Message = "Указанная директория не содержит файл для обработки!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории - """ _
                                                                  & FileSystem_FilePath & """"
            Exit Function
        Else
            ReDim tmp_Arr(1 To 1) As Variant
            tmp_Arr(1) = FileSystem_FilePath
            FileSystem_FilePath = tmp_Arr
        End If
    End If
    '```````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    FileSystem_File_Format = Right$(FileSystem_FilePath(1), _
                                Len(FileSystem_FilePath(1)) - InStrRev(FileSystem_FilePath(1), "."))
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case LCase$(FileSystem_File_Format)
        Case "xlsm", "xlsb", "xlam": Source_VBAProject = "\xl\vbaProject.bin"
        Case "docm", "dotm":         Source_VBAProject = "\word\vbaProject.bin"
        Case "pptm", "potm", "ppsm": Source_VBAProject = "\ppt\vbaProject.bin"

        Case "xls", "xla":           GoSub GS¦Hack_OldOffice_Excel:      GoTo GT¦Clearing_FileSystem
        Case "doc", "dot":           GoSub GS¦Hack_OldOffice_Word:       GoTo GT¦Clearing_FileSystem
        Case "ppt", "pot", "pps":    GoSub GS¦Hack_OldOffice_PowerPoint: GoTo GT¦Clearing_FileSystem

        Case "xlsx", "docx", "dotx", "pptx", "potx"
            Error_Message = "В файлах данного расширения (." & FileSystem_File_Format & ") защита VBAProject-a не реализуется!"
            Show_ErrorMessage_Immediate Error_Message, "Защита VBA_Project не обнаружена!": Exit Function
        Case Else
            Error_Message = "Данное расширение (." & FileSystem_File_Format & ") не поддерживается текущим функционалом!"
            Show_ErrorMessage_Immediate Error_Message, "Алгоритм не может быть выполнен для переданного файла!"
            Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Set Obj_FSO = CreateObject("Scripting.FileSystemObject")

    File_MetaData = Get_AllMetadata_File(CStr(FileSystem_FilePath(1)))
    File_PathFolder = Environ$("USERPROFILE") & FILE_SYSTEM_LOCAL_APP_DATA & FILE_SYSTEM_SAVE_PATH_UNPROTECT_VBPROJECT
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    If Dir(File_MetaData.MD_Folder_TempDirectory, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, File_MetaData.MD_Folder_TempDirectory, ByVal 0&)
    End If

    If Dir(File_PathFolder, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, File_PathFolder, ByVal 0&)
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    File_MetaData.MD_File_TempName = Replace(File_MetaData.MD_File_TempName, "." & _
                                             File_MetaData.MD_File_TypeName, ".zip")

    Set Obj_File = Obj_FSO.GetFile(FileSystem_FilePath(1)): Obj_File.Copy File_MetaData.MD_File_TempName
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Call ZIP_Archive_UnPack(CVar(File_MetaData.MD_Folder_TempDirectory), _
                            CVar(File_MetaData.MD_File_TempName))

    File_VBProject = File_MetaData.MD_Folder_TempDirectory & Source_VBAProject

    If Not Obj_FSO.FileExists(File_VBProject) Then
        Error_Message = "В выбранном файле отсутствует VBProject, который необходимо защитить - " & FileSystem_FilePath(1)
        Show_ErrorMessage_Immediate Error_Message, "Невозможно применить алгоритм защиты VBProject-а!"
        GoTo GT¦Clearing_FileSystem
    End If

    GoSub GS¦Hack_VBProject
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    ZIP_Archive_File = File_PathFolder & File_MetaData.MD_File_BaseName & ".zip"

    Call ZIP_Archive_Create(CVar(ZIP_Archive_File))
    Call ZIP_Archive_UnPack(CVar(ZIP_Archive_File), File_MetaData.MD_Folder_TempDirectory & "\")

    File_Name = File_MetaData.MD_File_BaseName & "." & File_MetaData.MD_File_TypeName
    FileSystem_FilePath = File_PathFolder & File_MetaData.MD_File_BaseName & "." & _
                                                                 File_MetaData.MD_File_TypeName

    If Obj_FSO.FileExists(FileSystem_FilePath) Then Obj_FSO.DeleteFile FileSystem_FilePath

    On Error Resume Next
    '```````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````'
GT¦ReName_ZIP_To_Excel:

    Obj_FSO.GetFile(ZIP_Archive_File).Name = File_Name
    If Not Err.Number = 0 Then Err = 0: GoTo GT¦ReName_ZIP_To_Excel

    Coll_FilesName.Add CStr(FileSystem_FilePath), FileSystem_FilePath

    FileSystem_Select_FilesInExplorer CStr(File_PathFolder), Coll_FilesName

    Audit_Unviewable = True

    On Error GoTo 0
'```````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````'
GT¦Clearing_FileSystem:

    On Error Resume Next
        Obj_FSO.DeleteFile File_MetaData.MD_File_TempName
        Obj_FSO.DeleteFolder File_MetaData.MD_Folder_TempDirectory
    On Error GoTo 0
'``````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````'
GT¦Clearing_Memory:

    Set Obj_File = Nothing:       Set Obj_FSO = Nothing
    Set Obj_FileOffice = Nothing: Set Obj_App = Nothing

    On Error Resume Next: Kill tmp_FilePath
    On Error GoTo 0:      Dir Environ$("TEMP")

    Exit Function
'```````````````````````````````````````````````````````'

'````````````````````````````````````````````````'
GS¦Hack_VBProject:

    FF = FreeFile()
    Open File_VBProject For Binary As #FF

    Buffer = Space$(LOF(FF)): Get #FF, , Buffer

    Buffer = Replace(Buffer, "CMG=""", "CMC=""")
    Buffer = Replace(Buffer, "DPB=""", "DPC=""")
    Buffer = Replace(Buffer, "GC=""", "CC=""")

    Seek #FF, 1: Put #FF, , Buffer: Close FF

    Return
'````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦Hack_OldOffice_Excel:

    If File_IsOpen(CStr(FileSystem_FilePath(1))) Then
        Error_Message = "Выбранный файл уже открыт (" & CStr(FileSystem_FilePath(1)) & _
                        ")! Требуется закрыть файл для корректной работы алгоритма!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
        Return
    End If

    Set Obj_App = CreateObject("Excel.Application")

    On Error Resume Next
    Set Obj_FileOffice = Obj_App.Workbooks.Open(FileSystem_FilePath(1), 0&, 1&)
    On Error GoTo 0

    If Obj_FileOffice Is Nothing Then
        Error_Message = "Проверьте файл (" & FileSystem_FilePath(1) & _
                                        "). Возможно он повреждён!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка открытия файла"
    Else
        Obj_App.DisplayAlerts = False
        tmp_FilePath = Convert_ExcelExtension(Obj_FileOffice, Environ$("TEMP") & "\" & _
                                              Obj_FileOffice.Name)
        If Len(tmp_FilePath) > 0 Then
            Flag_Result = Audit_Unviewable(tmp_FilePath)
            If Flag_Result Then Audit_Unviewable = True
            Obj_FileOffice.Close False: Kill tmp_FilePath
        End If
    End If

    Obj_App.Quit

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦Hack_OldOffice_Word:

    If File_IsOpen(CStr(FileSystem_FilePath(1))) Then
        Error_Message = "Выбранный файл уже открыт (" & CStr(FileSystem_FilePath(1)) & _
                        ")! Требуется закрыть файл для корректной работы алгоритма!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
        Return
    End If

    Set Obj_App = CreateObject("Word.Application")

    On Error Resume Next
    Set Obj_FileOffice = Obj_App.Documents.Open(FileSystem_FilePath(1))
    On Error GoTo 0

    If Obj_FileOffice Is Nothing Then
        Error_Message = "Проверьте файл (" & FileSystem_FilePath(1) & _
                                        "). Возможно он повреждён!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка открытия файла"
    Else
        Obj_App.DisplayAlerts = False
        tmp_FilePath = Convert_WordExtension(Obj_FileOffice, Environ$("TEMP") & "\" & _
                                             Obj_FileOffice.Name)
        If Len(tmp_FilePath) > 0 Then
            tmp_FilePath = Audit_Unviewable(tmp_FilePath)
            If Flag_Result Then Audit_Unviewable = True
            Obj_FileOffice.Close False: Kill tmp_FilePath
        End If
    End If

    Obj_App.Quit

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Hack_OldOffice_PowerPoint:

    If File_IsOpen(CStr(FileSystem_FilePath(1))) Then
        Error_Message = "Выбранный файл уже открыт (" & CStr(FileSystem_FilePath(1)) & _
                        ")! Требуется закрыть файл для корректной работы алгоритма!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
        Return
    End If

    Set Obj_App = CreateObject("PowerPoint.Application")

    On Error Resume Next
    Set Obj_FileOffice = Obj_App.Presentations.Open(FileSystem_FilePath(1))
    On Error GoTo 0

    If Obj_FileOffice Is Nothing Then
        Error_Message = "Проверьте файл (" & FileSystem_FilePath(1) & _
                                        "). Возможно он повреждён!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка открытия файла"
    Else
        Obj_App.DisplayAlerts = False
        tmp_FilePath = Convert_PowerPointExtension(Obj_FileOffice, Environ$("TEMP") & "\" & _
                                                   Obj_FileOffice.Name)
        If Len(tmp_FilePath) > 0 Then
            Flag_Result = Audit_Unviewable(tmp_FilePath)
            If Flag_Result Then Audit_Unviewable = True
            On Error Resume Next
            Obj_App.Presentations(1).Close False: On Error GoTo 0
        End If
    End If

    Obj_App.Quit

    Return
'````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================================'


'============================================================================================================================='
Public Function Audit_VBAProject( _
                Optional ByRef FileSystem_FilePath As Variant _
       ) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-----------------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````'
    Dim Obj_SDI As Object, Obj_VBComp  As Object, Obj_CodeModule As Object
    Dim Vector_stdModule_Code(1 To 36) As String, Obj_tmpVBComp  As Object
    Dim FileSystem_File_Format         As String, Code_Module   As Variant
    '`````````````````````````````````````````````````````````````````````````'
    Dim Flag_TimeLimit As Boolean
    Dim tmp_Arr()      As Variant, Time_Limit As Double
    Dim Error_Message  As String, Obj_WordApp  As Object, Obj_PPApp  As Object
    Dim Obj_ExlApp     As Object, Vector_stdModule_Code_RAM(1 To 22) As String
    '`````````````````````````````````````````````````````````````````````````'
    Const Std_Module_ID      As Long = 1, xl_Maximized       As Long = -4137
    Const pp_WindowMaximized As Long = 3, wd_WindowStateMaximize As Long = 1
    '`````````````````````````````````````````````````````````````````````````'

    '```````````````````````'
    Audit_VBAProject = False
    '```````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    If IsMissing(FileSystem_FilePath) Then
        FileSystem_FilePath = Get_List_Files_ToProcess(False, TID_FileSystem_Excel, _
                                                              TID_FileSystem_Word, _
                                                              TID_FileSystem_PowerPoint)
    Else
        If IsArray(FileSystem_FilePath) Or IsObject(FileSystem_FilePath) Then
            Error_Message = "Поддержка массивов и объектов не реализована!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории!"
            Exit Function
        End If

        If Not Get_Directory_Type(CStr(FileSystem_FilePath)) = DirType_File Then
            Error_Message = "Указанная директория не содержит файл для обработки!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории - """ _
                                                                  & FileSystem_FilePath & """"
            Exit Function
        Else
            ReDim tmp_Arr(1 To 1) As Variant
            tmp_Arr(1) = FileSystem_FilePath
            FileSystem_FilePath = tmp_Arr
        End If
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    If IsEmpty(FileSystem_FilePath(1)) Then Exit Function
    '````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    If File_IsOpen(CStr(FileSystem_FilePath(1))) Then
        Error_Message = "Выбранный файл уже открыт (" & CStr(FileSystem_FilePath(1)) & _
                        ")! Требуется закрыть файл для корректной работы алгоритма!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
        Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````` ``````````````````````````````````````````````````````````````````'
    FileSystem_File_Format = Right$(FileSystem_FilePath(1), _
                                Len(FileSystem_FilePath(1)) - InStrRev(FileSystem_FilePath(1), "."))
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case LCase$(FileSystem_File_Format)
        Case "xlsm", "xlsb", "xlam", "xls", "xla": GoSub GS¦Create_Process_Excel
        Case "doc", "docm", "dotm", "dot":         GoSub GS¦Create_Process_Word
        Case "ppt", "pptm", "potm":                GoSub GS¦Create_Process_PowerPoint

        Case "xlsx", "docx", "dotx", "pptx", "potx"
            Error_Message = "В файлах данного расширения (." & FileSystem_File_Format & ") защита VBAProject-a не реализуется!"
            Show_ErrorMessage_Immediate Error_Message, "Защита VBA_Project не обнаружена!"
        Case Else
            Error_Message = "Данное расширение (." & FileSystem_File_Format & ") не поддерживается текущим функционалом!"
            Show_ErrorMessage_Immediate Error_Message, "Алгоритм не может быть выполнен для переданного файла!"
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````'
GT¦Clearing_Memory:

    Set Obj_VBComp = Nothing: Set Obj_SDI = Nothing

    Exit Function
'``````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_Process_Excel:

    MacroSecurity_MacroParameters_Assign mso_Excel_1, Without_PrivilegesChanges, Access_Provided
    Set Obj_ExlApp = CreateObject("Excel.Application")

    Set Obj_SDI = Obj_ExlApp.Workbooks.Add
    Set Obj_VBComp = Obj_SDI.VBProject.VBComponents.Add(Std_Module_ID)
    Obj_VBComp.Name = "modKernel_Code"

    GoSub GS¦Fill_STD_Module
    If Flag_TimeLimit Then
        Obj_SDI.Workbooks.Close False: Obj_ExlApp.Quit
        Set Obj_ExlApp = Nothing: GoTo GT¦Clearing_Memory
    End If

    Obj_VBComp.CodeModule.AddFromString Code_Module

    With Obj_ExlApp
        .Run "BDFA5BB345245DDAF9EA6D2F742E28A9": .EnableEvents = False:
        .Workbooks.Open FileSystem_FilePath, 0&, 0&

        Select Case FileSystem_File_Format
            Case "xla", "xlam"
            Case Else: .Windows(2).WindowState = xl_Maximized
        End Select

        .Visible = True

        With .VBE.Windows("Immediate")
            .Visible = True: .SetFocus
        End With

    End With

    Audit_VBAProject = True

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_Process_Word:

    MacroSecurity_MacroParameters_Assign mso_Word_1, Without_PrivilegesChanges, Access_Provided
    Set Obj_WordApp = CreateObject("Word.Application")

    Set Obj_SDI = Obj_WordApp.Documents.Add
    Set Obj_VBComp = Obj_SDI.VBProject.VBComponents.Add(Std_Module_ID)
    Obj_VBComp.Name = "modKernel_Code"

    GoSub GS¦Fill_STD_Module_RAM
    If Flag_TimeLimit Then
        Obj_SDI.Documents.Close False: Obj_WordApp.Quit
        Set Obj_WordApp = Nothing: GoTo GT¦Clearing_Memory
    End If

    Obj_VBComp.CodeModule.AddFromString Code_Module

    With Obj_WordApp

        .Run "BDFA5BB345245DDAF9EA6D2F742E28A9"
        .Documents.Open FileSystem_FilePath(1)

        Select Case FileSystem_File_Format
            Case "dot", "dotm"
            Case Else: .Windows(2).WindowState = wd_WindowStateMaximize
        End Select

        .Visible = True

        With .VBE.Windows("Immediate")
            .Visible = True: .SetFocus
        End With

    End With

    Audit_VBAProject = True

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_Process_PowerPoint:

    MacroSecurity_MacroParameters_Assign mso_PowerPoint_1, Without_PrivilegesChanges, Access_Provided
    Set Obj_PPApp = CreateObject("PowerPoint.Application")

    Set Obj_SDI = Obj_PPApp.Presentations.Add
    Set Obj_VBComp = Obj_SDI.VBProject.VBComponents.Add(Std_Module_ID)
    Obj_VBComp.Name = "modKernel_Code"

    GoSub GS¦Fill_STD_Module_RAM
    If Flag_TimeLimit Then
        Obj_SDI.Presentations.Close False: Obj_PPApp.Quit
        Set Obj_PPApp = Nothing:     GoTo GT¦Clearing_Memory
    End If

    Obj_VBComp.CodeModule.AddFromString Code_Module

    With Obj_PPApp

        .Run "BDFA5BB345245DDAF9EA6D2F742E28A9"
        .Presentations.Open FileSystem_FilePath(1)

        Select Case FileSystem_File_Format
            Case "pot", "potm", "ppa", "ppam"
            Case Else: .Windows(2).WindowState = pp_WindowMaximized
        End Select

        With .VBE.Windows("Immediate")
            .Visible = True: .SetFocus
        End With

    End With

    Audit_VBAProject = True

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Fill_STD_Module:

    Time_Limit = Timer: Flag_TimeLimit = False

    On Error Resume Next
        While Obj_tmpVBComp Is Nothing
            Set Obj_tmpVBComp = Application.VBE.ActiveVBProject. _
                       VBComponents("VBD_Kit_Security"): DoEvents
            If Time_Limit + 5 < Timer Then
                Error_Message = "Невозможно получить ссылку на компонент! Проверьте привилегии!"
                Show_ErrorMessage_Immediate Error_Message, "Тайм-аут получения компонента!"
                Flag_TimeLimit = True: Return
            End If
        Wend
    On Error GoTo 0

    Set Obj_CodeModule = Obj_tmpVBComp.CodeModule

    Vector_stdModule_Code(1) = "'" & String$(79, "-") & "'"
    Vector_stdModule_Code(2) = Get_WinAPI_InModule(Obj_CodeModule, "Function VirtualProtect Lib") & vbCrLf
    Vector_stdModule_Code(3) = Get_WinAPI_InModule(Obj_CodeModule, "Function GetModuleHandleA Lib") & vbCrLf
    Vector_stdModule_Code(4) = Get_WinAPI_InModule(Obj_CodeModule, "Function GetProcAddress Lib") & vbCrLf
    Vector_stdModule_Code(5) = Get_WinAPI_InModule(Obj_CodeModule, "Sub MoveMemory Lib") & vbCrLf

    Vector_stdModule_Code(6) = Get_WinAPI_InModule(Obj_CodeModule, "Function VirtualAlloc Lib") & vbCrLf
    Vector_stdModule_Code(7) = Get_WinAPI_InModule(Obj_CodeModule, "Function VirtualFree Lib") & vbCrLf
    Vector_stdModule_Code(8) = Get_WinAPI_InModule(Obj_CodeModule, "Function CallDlgBxParam Lib") & vbCrLf

    Vector_stdModule_Code(9) = Get_WinAPI_InModule(Obj_CodeModule, "Function DialogBoxParam Lib")
    Vector_stdModule_Code(10) = "'" & String$(79, "-") & "'" & vbCrLf

    Vector_stdModule_Code(11) = Vector_stdModule_Code(11) & "'" & String$(41, "-") & "'"
    Vector_stdModule_Code(12) = Vector_stdModule_Code(12) & "Private Glb_Ptr_Func As LongPtr"
    Vector_stdModule_Code(13) = Vector_stdModule_Code(13) & "Private Glb_DefaultAddress    As LongPtr"
    Vector_stdModule_Code(14) = Vector_stdModule_Code(14) & "Private Glb_TrampolineAddress As LongPtr"
    Vector_stdModule_Code(15) = Vector_stdModule_Code(15) & "'" & String$(41, "-") & "'"
    Vector_stdModule_Code(16) = Vector_stdModule_Code(16) & "Private Glb_Hooking_Bytes(0 To 11) As Byte"
    Vector_stdModule_Code(17) = Vector_stdModule_Code(17) & "Private Glb_Default_Bytes(0 To 11) As Byte"
    Vector_stdModule_Code(18) = Vector_stdModule_Code(18) & "'" & String$(41, "-") & "'" & vbCrLf

    Vector_stdModule_Code(19) = Vector_stdModule_Code(19) & "'" & String$(69, "-") & "'"
    Vector_stdModule_Code(20) = Vector_stdModule_Code(20) & _
                                "Private Const CDP_HOOK_ACTIVE     As String = ""Hook_ActiveFlag"""
    Vector_stdModule_Code(21) = Vector_stdModule_Code(21) & _
                                "Private Const CDP_TRAMPOLINE_ADDR As String = ""Hook_TrampolineAddress"""
    Vector_stdModule_Code(22) = Vector_stdModule_Code(22) & _
                                "Private Const CDP_DEFAULT_ADDR    As String = ""UnHook_DefaultAddress"""
    Vector_stdModule_Code(23) = Vector_stdModule_Code(23) & _
                                "Private Const CDP_DEFAULT_BYTES   As String = ""UnHook_DefaultBytes"""
    Vector_stdModule_Code(24) = Vector_stdModule_Code(24) & "'" & String$(69, "-") & "'" & vbCrLf

    Vector_stdModule_Code(25) = Vector_stdModule_Code(25) & "'" & String$(92, "=") & "'"
    Vector_stdModule_Code(26) = Vector_stdModule_Code(26) & _
                               "Public Sub BDFA5BB345245DDAF9EA6D2F742E28A9(): " & _
                               "Call VBProject_Hooking_DialogBoxParam: End Sub "
    Vector_stdModule_Code(27) = Vector_stdModule_Code(27) & "'" & String$(92, "=") & "'" & vbCrLf

    Vector_stdModule_Code(28) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "VBProject_Hooking_DialogBoxParam") & vbCrLf
    Vector_stdModule_Code(29) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Hack_WinAPI_Intercept") & vbCrLf
    Vector_stdModule_Code(30) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Hack_DialogBoxParam") & vbCrLf
    Vector_stdModule_Code(31) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Get_PtrFunction") & vbCrLf
    Vector_stdModule_Code(32) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Release_Trampoline") & vbCrLf
    Vector_stdModule_Code(33) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Load_TrampolineAddress") & vbCrLf
    Vector_stdModule_Code(34) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Save_TrampolineState") & vbCrLf
    Vector_stdModule_Code(35) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Clear_TrampolineState") & vbCrLf
    Vector_stdModule_Code(36) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Show_ErrorMessage_Immediate") & vbCrLf

    Code_Module = Join(Vector_stdModule_Code, vbCrLf)
    Set Obj_tmpVBComp = Nothing: Set Obj_CodeModule = Nothing

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Fill_STD_Module_RAM:

    Time_Limit = Timer: Flag_TimeLimit = False

    On Error Resume Next
        While Obj_tmpVBComp Is Nothing
            Set Obj_tmpVBComp = Application.VBE.ActiveVBProject. _
                       VBComponents("VBD_Kit_Security"): DoEvents
            If Time_Limit + 5 < Timer Then
                Error_Message = "Невозможно получить ссылку на компонент! Проверьте привилегии!"
                Show_ErrorMessage_Immediate Error_Message, "Тайм-аут получения компонента!"
                Flag_TimeLimit = True: Return
            End If
        Wend
    On Error GoTo 0

    Set Obj_CodeModule = Obj_tmpVBComp.CodeModule

    Vector_stdModule_Code_RAM(1) = "'" & String$(79, "-") & "'"
    Vector_stdModule_Code_RAM(2) = Get_WinAPI_InModule(Obj_CodeModule, "Function VirtualProtect Lib") & vbCrLf
    Vector_stdModule_Code_RAM(3) = Get_WinAPI_InModule(Obj_CodeModule, "Function GetModuleHandleA Lib") & vbCrLf
    Vector_stdModule_Code_RAM(4) = Get_WinAPI_InModule(Obj_CodeModule, "Function GetProcAddress Lib") & vbCrLf
    Vector_stdModule_Code_RAM(5) = Get_WinAPI_InModule(Obj_CodeModule, "Sub MoveMemory Lib") & vbCrLf
    Vector_stdModule_Code_RAM(6) = Get_WinAPI_InModule(Obj_CodeModule, "Function DialogBoxParam Lib")
    Vector_stdModule_Code_RAM(7) = "'" & String$(79, "-") & "'" & vbCrLf

    Vector_stdModule_Code_RAM(8) = Vector_stdModule_Code_RAM(8) & "'" & String$(41, "-") & "'"
    Vector_stdModule_Code_RAM(9) = Vector_stdModule_Code_RAM(9) & "Private Glb_Ptr_Func As LongPtr"
    Vector_stdModule_Code_RAM(10) = Vector_stdModule_Code_RAM(10) & "'" & String$(41, "-") & "'"
    Vector_stdModule_Code_RAM(11) = Vector_stdModule_Code_RAM(11) & "Private Glb_Hooking_Bytes(0 To 11) As Byte"
    Vector_stdModule_Code_RAM(12) = Vector_stdModule_Code_RAM(12) & "Private Glb_Default_Bytes(0 To 11) As Byte"
    Vector_stdModule_Code_RAM(13) = Vector_stdModule_Code_RAM(13) & "'" & String$(41, "-") & "'" & vbCrLf
    Vector_stdModule_Code_RAM(14) = Vector_stdModule_Code_RAM(14) & "'" & String$(92, "=") & "'"
    Vector_stdModule_Code_RAM(15) = Vector_stdModule_Code_RAM(15) & _
                                    "Public Sub BDFA5BB345245DDAF9EA6D2F742E28A9(): " & _
                                    "Call VBProject_Hooking_DialogBoxParam_RAM: End Sub "
    Vector_stdModule_Code_RAM(16) = Vector_stdModule_Code_RAM(16) & "'" & String$(92, "=") & "'" & vbCrLf

    Vector_stdModule_Code_RAM(17) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "VBProject_Hooking_DialogBoxParam_RAM") & vbCrLf
    Vector_stdModule_Code_RAM(18) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "VBProject_Restore_DialogBoxParam") & vbCrLf
    Vector_stdModule_Code_RAM(19) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Hack_WinAPI_Intercept") & vbCrLf
    Vector_stdModule_Code_RAM(20) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Hack_DialogBoxParam_RAM") & vbCrLf
    Vector_stdModule_Code_RAM(21) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Get_PtrFunction") & vbCrLf
    Vector_stdModule_Code_RAM(22) = Get_Function_InModule(Obj_tmpVBComp, _
                                                     Obj_CodeModule, "Show_ErrorMessage_Immediate") & vbCrLf

    Code_Module = Join(Vector_stdModule_Code_RAM, vbCrLf)
    Set Obj_tmpVBComp = Nothing: Set Obj_CodeModule = Nothing

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================================='


'-----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-----------------------------------------------------------------------------------------------------------------------------'


'========================================================================================================================'
Public Function MacroProtection_ExcelObjectModel() As Variant
'------------------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox _
    "Данная функция (""MacroProtection_ExcelObjectModel"") находится в разработке!", vbInformation, "[DarkSec_Project]"
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================================'


'========================================================================================================================'
Public Function MacroProtection_OfficeCrypto() As Variant
'------------------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox _
    "Данная функция (""MacroProtection_OfficeCrypto"") не включена в текущую сборку модуля по инциативе разработчика!", _
                                                                             vbInformation, "[DarkSec_Project]"
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================================'


'=============================================================================================================================='
 Public Function MacroProtection_Unviewable( _
                Optional ByVal FileSystem_FilePath As Variant _
       ) As Variant
'------------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'------------------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````'
    Dim File_MetaData          As FileSystem_MetaData_File
    Dim File_PathFolder        As String, File_VBProject As String
    Dim FileSystem_File_Format As String, File_Name      As String
    '`````````````````````````````````````````````````````````````'
    Dim Buffer  As String, FF       As Integer
    Dim Obj_FSO As Object, Obj_File As Object
    '`````````````````````````````````````````````````````````````'
    Dim ZIP_Archive_File  As String
    Dim Source_VBAProject As String
    Dim Coll_FilesName    As New Collection
    Dim tmp_Arr           As Variant, Error_Message As String
    Dim Number_Changes    As Integer
    '`````````````````````````````````````````````````````````````'
    Dim Change_PositionIndex As Long
    Dim File_Length          As Long
    Dim File_Content()       As Byte
    '`````````````````````````````````````````````````````````````'
    Dim Vector_RootEntry()         As Variant
    Dim Vector_VersionCompatible() As Variant
    Dim Vector_GlobalChanges()     As Variant
    '`````````````````````````````````````````````````````````````'
    Dim i As Long, L As Long, K As Long, Pos As Long
    Dim Is_Match  As Boolean, Is_Next     As Boolean
    '`````````````````````````````````````````````````````````````'
    Dim RootEntry_LastID         As Long
    Dim VersionCompatible_LastID As Long
    Dim GlobalChanges_LastID     As Long
    '`````````````````````````````````````````````````````````````'
    Dim Changes() As Change_ContextVBProject, Change_Value As Byte
    '`````````````````````````````````````````````````````````````'
    Dim BRead_Byte(1 To 39) As Byte, CRead_Byte(1 To 4) As Byte
    '`````````````````````````````````````````````````````````````'

    '````````````````````````````````'
    MacroProtection_Unviewable = False
    '````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If IsMissing(FileSystem_FilePath) Then
        FileSystem_FilePath = Get_List_Files_ToProcess(False, TID_FileSystem_Excel, _
                                                              TID_FileSystem_Word, _
                                                              TID_FileSystem_PowerPoint)
        If Len(FileSystem_FilePath(1)) = 0 Then Exit Function
    Else
        If IsArray(FileSystem_FilePath) Or IsObject(FileSystem_FilePath) Then
            Error_Message = "Поддержка массивов и объектов не реализована!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории!"
            Exit Function
        End If

        If Not Get_Directory_Type(CStr(FileSystem_FilePath)) = DirType_File Then
            Error_Message = "Указанная директория не содержит файл для обработки!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории - """ _
                                                                  & FileSystem_FilePath & """"
            Exit Function
        Else
            ReDim tmp_Arr(1 To 1) As Variant
            tmp_Arr(1) = FileSystem_FilePath
            FileSystem_FilePath = tmp_Arr
        End If
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    FileSystem_File_Format = Right$(FileSystem_FilePath(1), _
                                Len(FileSystem_FilePath(1)) - InStrRev(FileSystem_FilePath(1), "."))
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case LCase$(FileSystem_File_Format)
        Case "xlsm", "xlsb", "xlam": Source_VBAProject = "\xl\vbaProject.bin"
        Case "docm", "dotm":         Source_VBAProject = "\word\vbaProject.bin"
        Case "pptm", "potm":         Source_VBAProject = "\ppt\vbaProject.bin"

        Case "xlsx", "docx", "dotx", "pptx", "potx"
            Error_Message = "В файлах данного расширения (." & FileSystem_File_Format & ") защита VBAProject-a не реализуется!"
            Show_ErrorMessage_Immediate Error_Message, "Защита VBA_Project не требуется!": Exit Function
        Case Else
            Error_Message = "Данное расширение (." & FileSystem_File_Format & ") не поддерживается текущим функционалом!"
            Show_ErrorMessage_Immediate Error_Message, "Алгоритм не может быть выполнен для переданного файла!"
            Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Set Obj_FSO = CreateObject("Scripting.FileSystemObject")
    File_MetaData = Get_AllMetadata_File(CStr(FileSystem_FilePath(1)))
    File_PathFolder = Environ$("USERPROFILE") & FILE_SYSTEM_LOCAL_APP_DATA & FILE_SYSTEM_SAVE_PATH_PROTECTED_VBPROJECT
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    If Dir(File_MetaData.MD_Folder_TempDirectory, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, File_MetaData.MD_Folder_TempDirectory, ByVal 0&)
    End If

    If Dir(File_PathFolder, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(ByVal 0&, File_PathFolder, ByVal 0&)
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    File_MetaData.MD_File_TempName = Replace(File_MetaData.MD_File_TempName, "." & _
                                             File_MetaData.MD_File_TypeName, ".zip")

    Set Obj_File = Obj_FSO.GetFile(FileSystem_FilePath(1)): Obj_File.Copy File_MetaData.MD_File_TempName
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Call ZIP_Archive_UnPack(CVar(File_MetaData.MD_Folder_TempDirectory), _
                            CVar(File_MetaData.MD_File_TempName))

    File_VBProject = File_MetaData.MD_Folder_TempDirectory & Source_VBAProject

    If Not Obj_FSO.FileExists(File_VBProject) Then
        Error_Message = "В выбранном файле отсутствует VBProject, который необходимо защитить - " & FileSystem_FilePath(1)
        Show_ErrorMessage_Immediate Error_Message, "Невозможно применить алгоритм защиты VBProject-а!"
        GoTo GT¦Clearing_FileSystem
    End If

    File_Length = FileLen(File_VBProject)
    FF = FreeFile: Open File_VBProject For Binary As #FF Len = 1: GoSub GS¦Set_Unviewable: Close #FF

    If Change_PositionIndex = 0 Then
        Error_Message = "В выбранном файле не удалось установить защиту VBProject-a! " & _
                        "Возможно она уже присутствует или файл имеет нестандартную структуру VBProject-а!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно применить алгоритм защиты VBProject-а!"
        GoTo GT¦Clearing_FileSystem
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    File_MetaData.MD_File_BaseName = Replace(File_MetaData.MD_File_BaseName, "_Viewable", _
                                                                              vbNullString)

    File_MetaData.MD_File_BaseName = Replace(File_MetaData.MD_File_BaseName, "_Unviewable", _
                                                                              vbNullString)

    ZIP_Archive_File = File_PathFolder & File_MetaData.MD_File_BaseName & ".zip"

    Call ZIP_Archive_Create(CVar(ZIP_Archive_File))
    Call ZIP_Archive_UnPack(CVar(ZIP_Archive_File), File_MetaData.MD_Folder_TempDirectory & "\")

    File_Name = File_MetaData.MD_File_BaseName & "." & File_MetaData.MD_File_TypeName
    FileSystem_FilePath = File_PathFolder & File_MetaData.MD_File_BaseName & "." & _
                                                                 File_MetaData.MD_File_TypeName

    If Obj_FSO.FileExists(FileSystem_FilePath) Then Obj_FSO.DeleteFile FileSystem_FilePath

    Obj_FSO.GetFile(ZIP_Archive_File).Name = File_Name
    Coll_FilesName.Add CStr(FileSystem_FilePath), FileSystem_FilePath

    FileSystem_Select_FilesInExplorer CStr(File_PathFolder), Coll_FilesName
    MacroProtection_Unviewable = FileSystem_FilePath
    '```````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````'
GT¦Clearing_FileSystem:

    On Error Resume Next
        Obj_FSO.DeleteFile File_MetaData.MD_File_TempName
        Obj_FSO.DeleteFolder File_MetaData.MD_Folder_TempDirectory
    On Error GoTo 0

    Dir Environ$("TEMP")

    Exit Function
'``````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````'
GS¦Set_Unviewable:

    Vector_RootEntry = Array(&H52, &H0, &H6F, &H0, _
              &H6F, &H0, &H74, &H0, _
              &H20, &H0, &H45, &H0, _
              &H6E, &H0, &H74, &H0, _
              &H72, &H0, &H79, &H0, _
              &H0, &H0, &H0, &H0, _
              &H0, &H0, &H0, &H0, _
              &H0, &H0, &H0, &H0, _
              &H0, &H0, &H0, &H0, _
              &H0, &H0, &H0 _
              )

    Vector_VersionCompatible = Array( _
              &H56, &H65, &H72, &H73, _
              &H69, &H6F, &H6E, &H43, _
              &H6F, &H6D, &H70, &H61, _
              &H74, &H69, &H62, &H6C, _
              &H65, &H33, &H32, &H3D, _
              &H22, &H33, &H39, &H33, _
              &H32, &H32, &H32, &H30, _
              &H30, &H30, &H22, &HD, _
              &HA, &H43, &H4D, &H47, _
              &H3D, &H22 _
              )

    Vector_GlobalChanges = Array( _
              &H47, &H43, &H3D, &H22 _
              )

    RootEntry_LastID = UBound(Vector_RootEntry()) + 1
    VersionCompatible_LastID = UBound(Vector_VersionCompatible()) + 1
    GlobalChanges_LastID = UBound(Vector_GlobalChanges()) + 1

    ReDim File_Content(1 To File_Length)
    Get #FF, , File_Content: Change_PositionIndex = 0

    For i = 1 To File_Length - RootEntry_LastID
        Is_Match = True
        For K = 1 To RootEntry_LastID
            If Vector_RootEntry(K - 1) <> File_Content(i + K - 1) Then
                Is_Match = False: Exit For
            End If
        Next

        If Is_Match Then
            For L = 1 To 5
                File_Content(i + 235 + L - 1) = CByte(255 * Rnd())
            Next
            Change_PositionIndex = i: Exit For
        End If
    Next

    Put #FF, , File_Content

    If Not Change_PositionIndex = 0 Then
        Change_PositionIndex = 0: Number_Changes = 0

        For i = 1 To File_Length - VersionCompatible_LastID
            Get #FF, i, BRead_Byte: Is_Match = True
            For K = 1 To VersionCompatible_LastID
                If Vector_VersionCompatible(K - 1) <> BRead_Byte(K) Then
                    Is_Match = False
                    Exit For
                End If
            Next

            If Is_Match Then
                Pos = i + VersionCompatible_LastID: Is_Next = True

                Do While Is_Next
                    Get #FF, Pos, Change_Value
                    If Change_Value = &H22 Then
                        Is_Next = False
                    Else
                        Number_Changes = Number_Changes + 1
                        ReDim Preserve Changes(1 To Number_Changes)
                        Changes(Number_Changes).Position = Pos
                        Changes(Number_Changes).Value = &H46
                    End If
                    Pos = Pos + 1
                Loop

                Exit For
            End If
        Next

        If Is_Match Then
            For K = Pos To File_Length - UBound(Vector_GlobalChanges) + 1
                Get #FF, K, CRead_Byte
                Is_Match = True

                For L = 1 To GlobalChanges_LastID
                    If Vector_GlobalChanges(L - 1) <> CRead_Byte(L) Then
                        Is_Match = False
                        Exit For
                    End If
                Next

                If Is_Match Then
                    Pos = K + GlobalChanges_LastID
                    Is_Next = True

                    Do While Is_Next
                        Get #FF, Pos, Change_Value
                        If Change_Value = &H22 Then
                            Is_Next = False
                        Else
                            Number_Changes = Number_Changes + 1
                            ReDim Preserve Changes(1 To Number_Changes)
                            Changes(Number_Changes).Position = Pos
                            Changes(Number_Changes).Value = &H46
                        End If
                        Pos = Pos + 1
                    Loop

                    Change_PositionIndex = Pos
                    Exit For
                End If
            Next
        End If

        If Not Change_PositionIndex = 0 Then
            For i = 1 To Number_Changes
                Put #FF, Changes(i).Position, Changes(i).Value
            Next
        End If
    End If

    Return
'`````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------------'
End Function
'=============================================================================================================================='


'==================================================================================================================='
Public Function MacroProtection_VBAProject( _
                Optional ByVal FileSystem_FilePath As Variant, _
                Optional ByVal VB_Password As String = "PswrD" _
       ) As Variant
'-------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-------------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````'
    Dim Handle_WinDlg As LongPtr, Handle_SysTab As LongPtr
    Dim Obj_ExlApp As Object, Obj_WB As Object, VBProject_Name As String
    Dim File_PathFolder As String, File_NewName As String, File_Path As String
    Dim Coll_FilesName As New Collection, Error_Message As String
    Dim Flag_ProtectionRequired As Boolean, Inx_Pos As Long
    Dim Obj_VBProject As Object
    '````````````````````````````````````````````````````````````````````````````'
    Const TCM_SETCURSEL = &H130C, TCM_SETCURFOCUS = &H1330, EM_SETMODIFY = &HB9
    Const BST_CHECKED = &H1, WM_SETTEXT = &HC, GW_CHILD = 5
    Const BM_SETCHECK = &HF1, BM_GETCHECK = &HF0, BM_CLICK = &HF5
    Const VB_Protection_None As Long = &H0, xl_OpenXMLAddIn = 55&, xl_AddIn = 18&
    '````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````'
    MacroProtection_VBAProject = False
    '````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    If IsMissing(FileSystem_FilePath) Then
        FileSystem_FilePath = Get_List_Files_ToProcess(False, TID_FileSystem_Excel)
    Else
        If IsArray(FileSystem_FilePath) Or IsObject(FileSystem_FilePath) Then
            Error_Message = "Поддержка массивов и объектов не реализована!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории!"
            Exit Function
        End If

        If Not Get_Directory_Type(CStr(FileSystem_FilePath)) = DirType_File Then
            Error_Message = "Указанная директория не содержит файл для обработки - """"" _
                                                            & FileSystem_FilePath & """"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка обработки директории/файла"
            Exit Function
        Else
            ReDim tmp_Arr(1 To 1) As Variant
            tmp_Arr(1) = FileSystem_FilePath
            FileSystem_FilePath = tmp_Arr
        End If
    End If
    '`````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    If IsEmpty(FileSystem_FilePath(1)) Then Exit Function
    '````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    FileSystem_FilePath = FileSystem_FilePath(1)
    File_NewName = CreateObject("Scripting.FileSystemObject").GetFileName(FileSystem_FilePath)
    File_Path = Mid$(FileSystem_FilePath, 1, InStr(FileSystem_FilePath, File_NewName) - 1)
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Select Case True

        Case File_NewName Like "*.xlsx"
            Error_Message = "В файлах данного расширения (." & _
                             File_NewName & ") защита VBAProject-a не реализуется!"
            Show_ErrorMessage_Immediate Error_Message, "Защита VBA_Project не требуется!"

        Case File_NewName Like "*.xls*" Or File_NewName Like "*.xla*"

            On Error Resume Next

            If Len(ThisWorkbook.VBProject.Name) = 0 Then
                Call MacroSecurity_MacroParameters_Assign( _
                            mso_Excel_1, Without_PrivilegesChanges, Access_Provided, False _
                     )
            End If

            On Error GoTo 0

            Set Obj_ExlApp = CreateObject("Excel.Application")
            Obj_ExlApp.EnableEvents = False

            If File_IsOpen(CStr(FileSystem_FilePath)) Then
                Error_Message = "Выбранный файл уже открыт (" & CStr(FileSystem_FilePath) & _
                                ")! Требуется закрыть файл для корректной работы алгоритма!"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
            Else
                On Error Resume Next
                Set Obj_WB = Obj_ExlApp.Workbooks.Open(FileSystem_FilePath, False, False)
                On Error GoTo 0

                If Obj_WB Is Nothing Then
                    Error_Message = "Проверьте файл (" & FileSystem_FilePath & _
                                    "). Возможно он повреждён!"
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка открытия файла"
                    GoTo GT¦Clearing_Memory
                End If

                GoSub GS¦Set_Password: If Flag_ProtectionRequired Then GoSub GS¦Create_BackUps
            End If

        Case Else
            Error_Message = "Функционал не поддерживает текущий формат файлов!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка выполнения алгоритма!"
            On Error GoTo 0: Exit Function

    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````'
GT¦Clearing_Memory:

    On Error Resume Next

    Obj_WB.Close False:  Set Obj_WB = Nothing
    Obj_ExlApp.Quit: Set Obj_ExlApp = Nothing

    Dir Environ$("TEMP"): On Error GoTo 0

    Exit Function
'`````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_BackUps:

    File_PathFolder = Environ$("USERPROFILE") & FILE_SYSTEM_LOCAL_APP_DATA & _
                                                FILE_SYSTEM_SAVE_PATH_PROTECTED_VBPROJECT

    If Dir(File_PathFolder, vbDirectory) = vbNullString Then
        Call SHCreateDirectoryEx(Application.hWnd, File_PathFolder, 0&)
    End If

    FileSystem_FilePath = File_PathFolder & File_NewName

    If File_IsOpen(CStr(FileSystem_FilePath)) Then
        Inx_Pos = InStrRev(File_NewName, ".")
        File_NewName = Mid$(File_NewName, 1&, Inx_Pos - 1&) & "_(" & _
                            CreateObject("Scripting.FileSystemObject").GetTempName & ")" & _
                       Mid$(File_NewName, Inx_Pos)
        FileSystem_FilePath = File_PathFolder & File_NewName
    End If

    On Error Resume Next

    With Obj_ExlApp
        .DisplayAlerts = False
            If Right$(FileSystem_FilePath, 4) = "xlam" Then
                Obj_WB.SaveAs FileName:=FileSystem_FilePath, FileFormat:=xl_OpenXMLAddIn
            ElseIf Right$(FileSystem_FilePath, 3) = "xla" Then
                Obj_WB.SaveAs FileName:=FileSystem_FilePath, FileFormat:=xl_AddIn
            Else
                Obj_WB.SaveCopyAs FileName:=FileSystem_FilePath
            End If
        .DisplayAlerts = True
    End With

    If Err.Number = 0 Then
        Coll_FilesName.Add FileSystem_FilePath, FileSystem_FilePath
        FileSystem_Select_FilesInExplorer CStr(File_PathFolder), Coll_FilesName
        MacroProtection_VBAProject = Coll_FilesName
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Set_Password:

    If Obj_ExlApp.VBE.ActiveVBProject.Protection = VB_Protection_None Then
        VBProject_Name = Obj_ExlApp.VBE.ActiveVBProject.Name
    Else
        Error_Message = "Выбранный файл уже содержит защиту проекта VBA - """ & FileSystem_FilePath & """"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка установки пароля на проект VBA"
        GoTo GT¦Clearing_Memory
    End If

    Flag_ProtectionRequired = Is_VBProjectProtectionRequired(Obj_WB)

    If Flag_ProtectionRequired Then

        Obj_ExlApp.VBE.CommandBars(1).FindControl(ID:=2578, Recursive:=True).Execute
        Handle_WinDlg = FindWindow(vbNullString, VBProject_Name & " - Project Properties")

        If Not Handle_WinDlg = 0 Then
            Handle_SysTab = FindWindowEx(Handle_WinDlg, 0, "SysTabControl32", vbNullString)
        Else
            GoTo GT¦Clearing_Memory
        End If

        Call SendMessage(Handle_SysTab, TCM_SETCURFOCUS, 1, 0): Call SendMessage(Handle_SysTab, TCM_SETCURSEL, 1, 0)

        If SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1557), BM_GETCHECK, 0, 0) = 0 Then
            Call SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1557), BM_SETCHECK, BST_CHECKED, 0)
            Call SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1555), WM_SETTEXT, 0, VB_Password)
            Call SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1555), EM_SETMODIFY, True, 0)
            Call SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1556), WM_SETTEXT, 0, VB_Password)
            Call SendMessage(GetDlgItem(GetNextWindow(Handle_WinDlg, GW_CHILD), &H1556), EM_SETMODIFY, True, 0)
        End If

        DoEvents: Call SendMessage(GetDlgItem(Handle_WinDlg, &H1), BM_CLICK, 0, 0): DoEvents

    Else
        Error_Message = "В данном файле (" & FileSystem_FilePath & _
                        ") нет кода, который требуется защитить!": Err.Clear
        Show_ErrorMessage_Immediate Error_Message, "Защита VBA_Project не требуется!"
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'-------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-------------------------------------------------------------------------------------------------------------------'


'================================================================================================='
Public Function MacroSecurity_ExternalContent_Assign( _
                ByVal Type_ExternalContent As MSExcel_Type_ExternalContent, _
                ByVal Setting_ExternalContent As MSExcel_Setting_ExternalContent _
       ) As Boolean
'-------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, API_lpData As Long, Registry_Section As String
    Dim Section_Data As Registry_SectionData, ProtectedView_State As Long
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    MacroSecurity_ExternalContent_Assign = False
    '```````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1) & _
                                              REGISTRY_SECTION_EXCEL_SECURITY
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_ALL_ACCESS)
    Handle = Section_Data.Handle: If Handle = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Call VBD_Kit_Security.MacroSecurity_ExternalContent_Fetch(Type_ExternalContent, True)

    Select Case Type_ExternalContent
        Case mse_WorkbookLinkWarnings:   GoSub GS¦Change_Registry_WorkbookLinkWarnings
        Case mse_DataConnectionWarnings: GoSub GS¦Change_Registry_DataConnectionWarnings
    End Select

    Call RegCloseKey(Handle): Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_DataConnectionWarnings:

    Call RegQueryValueEx(Handle, "DataConnectionWarnings", 0&, _
                         REG_DWORD, API_lpData, 4&)

    If API_lpData <> Setting_ExternalContent Then
        MacroSecurity_ExternalContent_Assign = RegSetValueEx(Handle, "DataConnectionWarnings", 0&, _
                                                      REG_DWORD, Setting_ExternalContent, 4&) = 0&
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_WorkbookLinkWarnings:

    Call RegQueryValueEx(Handle, "WorkbookLinkWarnings", 0&, _
                         REG_DWORD, API_lpData, 4&)

    If API_lpData <> Setting_ExternalContent Then
        MacroSecurity_ExternalContent_Assign = RegSetValueEx(Handle, "WorkbookLinkWarnings", 0&, _
                                                      REG_DWORD, Setting_ExternalContent, 4&) = 0&
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'===================================================================================================================='
Public Function MacroSecurity_ExternalContent_Fetch( _
                ByVal Type_ExternalContent As MSExcel_Type_ExternalContent, _
                Optional ByVal Status_Code As Boolean = False _
       ) As Variant
'--------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'--------------------------------------------------------------------------------------------------------------------'
    
    '``````````````````````````````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, WinAPI_Result As Long, API_lpData   As Long, API_lpSubKey As String
    Dim Section_Data As Registry_SectionData, Registry_Section As String, Sub_Handle As LongPtr
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````'
    MacroSecurity_ExternalContent_Fetch = False
    '``````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1) & _
                                              REGISTRY_SECTION_EXCEL_SECURITY
    '```````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_READ)
    If Len(Section_Data.Section_Path) = 0 Then Exit Function
    Handle = Section_Data.Handle: If Handle = 0 Then GoSub GS¦Create_DefaultKey
    '`````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_ExternalContent

        Case mse_DataConnectionWarnings
            WinAPI_Result = RegQueryValueEx(Handle, "DataConnectionWarnings", 0&, _
                                            REG_DWORD, API_lpData, 4&)
            If WinAPI_Result = ERROR_FILE_NOT_FOUND Then
                API_lpSubKey = "DataConnectionWarnings"
                GoSub GS¦Create_DefaultKey: API_lpData = 0&
            End If

        Case mse_WorkbookLinkWarnings
            WinAPI_Result = RegQueryValueEx(Handle, "WorkbookLinkWarnings", 0&, _
                                            REG_DWORD, API_lpData, 4&)
            If WinAPI_Result = ERROR_FILE_NOT_FOUND Then
                API_lpSubKey = "WorkbookLinkWarnings"
                GoSub GS¦Create_DefaultKey: API_lpData = 2&
            End If

    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Status_Code Then
        Select Case API_lpData
            Case 0
                If Type_ExternalContent = mse_DataConnectionWarnings Then
                    MacroSecurity_ExternalContent_Fetch = "Все подключения включены (не рекомендуется)"
                Else
                    MacroSecurity_ExternalContent_Fetch = "Все обновления связей в книге включены (не рекомендуется)"
                End If

            Case 1
                If Type_ExternalContent = mse_DataConnectionWarnings Then
                    MacroSecurity_ExternalContent_Fetch = "Запрос на подключение к данным"
                Else
                    MacroSecurity_ExternalContent_Fetch = "Запрос на автоматическое обновление связей в книге"
                End If

            Case 2
                If Type_ExternalContent = mse_DataConnectionWarnings Then
                    MacroSecurity_ExternalContent_Fetch = "Все подключения отключены"
                Else
                    MacroSecurity_ExternalContent_Fetch = "Все обновления связей в книге отключены"
                End If
        End Select
    Else
        MacroSecurity_ExternalContent_Fetch = API_lpData
    End If

    Call RegCloseKey(Handle)

    Exit Function
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_DefaultKey:

    WinAPI_Result = RegOpenKeyEx(HKEY_CURRENT_USER, Registry_Section, 0&, KEY_ALL_ACCESS, Handle)
    If Not WinAPI_Result = 0 Then Exit Function

    Call RegSetValueEx(Handle, CStr(API_lpSubKey), 0&, REG_DWORD, 2&, 4&)

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------------'
End Function
'===================================================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================================'
Public Function MacroSecurity_MacroParameters_Assign( _
                ByVal Type_Document As MSOffice_Type_Document_Group_1, _
                Optional ByVal Macro_Privileges As Security_ManagementCenter_Privileges = Without_PrivilegesChanges, _
                Optional ByVal Macro_AccessObjectModel As Security_ManagementCenter_AccessObjectModel = Without_AOMChanges, _
                Optional ByVal Macro_ForcedChange As Boolean = True _
       ) As Boolean
'----------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'----------------------------------------------------------------------------------------------------------------------------'

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

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document

        Case MSOffice_Type_Document_Group_1.mso_Access_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_ACCESS_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Excel_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Outlook_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_OUTLOOK_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_PowerPoint_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Word_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_WORD_SECURITY

    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_ALL_ACCESS)
    Handle = Section_Data.Handle: If Handle = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_Document

        Case mso_Access_1
             If Macro_Privileges = Without_PrivilegesChanges Then Exit Function
             GoSub GS¦Change_Registry_VBAWarnings

        Case mso_Excel_1
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then Native_Process = True

        Case mso_PowerPoint_1
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_PowerPoint Then Native_Process = True

        Case mso_Word_1
             If Macro_Privileges <> Without_PrivilegesChanges Then GoSub GS¦Change_Registry_VBAWarnings
             If Macro_AccessObjectModel <> Without_AOMChanges Then GoSub GS¦Change_Registry_AccessVBOM
             If Glb_MSOffice_Type_Application = Type_MSOffice_Word Then Native_Process = True

        Case mso_Outlook_1
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

'``````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_VBAWarnings:

    Call RegQueryValueEx(Handle, "VBAWarnings", 0&, REG_DWORD, API_lpData, 4&)

    If API_lpData <> Macro_Privileges Then
        MacroSecurity_MacroParameters_Assign = RegSetValueEx(Handle, "VBAWarnings", 0&, _
                                             REG_DWORD, Macro_Privileges, 4&) = 0&
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````'

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

'----------------------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================================'


'======================================================================================================================='
Public Function MacroSecurity_MacroParameters_Fetch( _
                ByVal Type_Document As MSOffice_Type_Document_Group_1, _
                ByVal Macro_Param As Security_SettingMacro, _
                Optional ByVal Status_Code As Boolean = False _
       ) As Variant
'-----------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-----------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, API_lpData As Long, Error_Message   As String
    Dim Section_Data As Registry_SectionData, Registry_Section As String
    '```````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````'
    MacroSecurity_MacroParameters_Fetch = False
    '``````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document

        Case MSOffice_Type_Document_Group_1.mso_Access_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_ACCESS_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Excel_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Outlook_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_OUTLOOK_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_PowerPoint_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY

        Case MSOffice_Type_Document_Group_1.mso_Word_1
             Registry_Section = Registry_Section & REGISTRY_SECTION_WORD_SECURITY

    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_READ)
    Handle = Section_Data.Handle: If Handle = 0 Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_Document

        Case mso_Excel_1, mso_Word_1, mso_PowerPoint_1
            If Macro_Param = Privileges Then GoTo GT¦Get_Registry_VBAWarnings
            If Macro_Param = AccessObjectModel Then GoTo GT¦Get_Registry_AccessVBOM

        Case mso_Access_1
            If Macro_Param = Privileges Then GoTo GT¦Get_Registry_VBAWarnings
            If Macro_Param = AccessObjectModel Then
                If Status_Code Then
                    MacroSecurity_MacroParameters_Fetch = -1
                Else
                    MacroSecurity_MacroParameters_Fetch = "Данный параметр безопасности отсутствует в MS Access"
                End If
            End If

        Case mso_Outlook_1
            If Macro_Param = Privileges Then GoTo GT¦Get_Registry_Level
            If Macro_Param = AccessObjectModel Then GoTo GT¦Get_Registry_DontTrustInstalledFiles

    End Select

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Get_Registry_VBAWarnings:

    Call RegQueryValueEx(Handle, "VBAWarnings", 0&, REG_DWORD, API_lpData, 4&)

    If Status_Code Then
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_WithNotification
            Case 1: MacroSecurity_MacroParameters_Fetch = Enable_AllMacros_NotRecommended
            Case 2: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_WithNotification
            Case 3: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_ExceptDigitallySignedMacros
            Case 4: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros
        End Select
    Else
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (с уведомлением)"
            Case 1: MacroSecurity_MacroParameters_Fetch = "Все макросы разрешены (не рекомендуется)"
            Case 2: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (с уведомлением)"
            Case 3: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (кроме макросов с цифровой подписью)"
            Case 4: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (без уведомления)"
        End Select
    End If

    Call RegCloseKey(Handle)

    Exit Function
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Get_Registry_AccessVBOM:

    Call RegQueryValueEx(Handle, "AccessVBOM", 0&, REG_DWORD, API_lpData, 4&)

    If Status_Code Then
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = Access_Denied
            Case 1: MacroSecurity_MacroParameters_Fetch = Access_Provided
        End Select
    Else
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = "Доступ к объектной модели запрещён"
            Case 1: MacroSecurity_MacroParameters_Fetch = "Доступ к объектной модели разрешен"
        End Select
    End If

    Call RegCloseKey(Handle)

    Exit Function
'`````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Get_Registry_Level:

    Call RegQueryValueEx(Handle, "Level", 0&, REG_DWORD, API_lpData, 4&)

    If Status_Code Then
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_WithNotification
            Case 1: MacroSecurity_MacroParameters_Fetch = Enable_AllMacros_NotRecommended
            Case 2: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_WithNotification
            Case 3: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros_ExceptDigitallySignedMacros
            Case 4: MacroSecurity_MacroParameters_Fetch = Disable_AllMacros
        End Select
    Else
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (кроме макросов с цифровой подписью)"
            Case 1: MacroSecurity_MacroParameters_Fetch = "Все макросы разрешены (не рекомендуется)"
            Case 2: MacroSecurity_MacroParameters_Fetch = "Уведомления для всех макросов"
            Case 3: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены (кроме макросов с цифровой подписью)"
            Case 4: MacroSecurity_MacroParameters_Fetch = "Все макросы отключены без уведомления"
        End Select
    End If

    Call RegCloseKey(Handle)

    Exit Function
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Get_Registry_DontTrustInstalledFiles:

    Call RegQueryValueEx(Handle, "DontTrustInstalledFiles", 0&, REG_DWORD, API_lpData, 4&)

    If Status_Code Then
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = Access_Denied
            Case 1: MacroSecurity_MacroParameters_Fetch = Access_Provided
        End Select
    Else
        Select Case API_lpData
            Case 0: MacroSecurity_MacroParameters_Fetch = "Параметры безопасности макросов не применяются к надстройкам"
            Case 1: MacroSecurity_MacroParameters_Fetch = "Параметры безопасности макросов применяются к надстройкам"
        End Select
    End If

    Call RegCloseKey(Handle)

    Exit Function
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================================='


'========================================================================================================'
Public Function MacroSecurity_ProtectedView_Assign( _
                ByVal Type_Document As MSOffice_Type_Document_Group_3, _
                ByVal Type_ProtectedView As MSOffice_Type_ProtectedView, _
                Optional ByVal Activate_ProtectedView As Boolean = False _
       ) As Boolean
'--------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'--------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, API_lpData As Long, Registry_Section As String
    Dim Section_Data As Registry_SectionData, ProtectedView_State As Long
    '````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    MacroSecurity_ProtectedView_Assign = False
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Call MacroSecurity_ProtectedView_Fetch(Type_Document, Type_ProtectedView, True)

    Select Case Type_Document

        Case mso_Excel_3
             Registry_Section = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY & "ProtectedView"

        Case mso_PowerPoint_3
             Registry_Section = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY & "ProtectedView"

        Case mso_Word_3
             Registry_Section = Registry_Section & REGISTRY_SECTION_WORD_SECURITY & "ProtectedView"

    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_ALL_ACCESS)
    Handle = Section_Data.Handle: If Handle = 0 Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    ProtectedView_State = IIf(Activate_ProtectedView, 0, 1)

    Select Case Type_ProtectedView
        Case mso_AttachmentsInPV:     GoSub GS¦Change_Registry_DisableAttachmentsInPV
        Case mso_InternetFilesInPV:   GoSub GS¦Change_Registry_DisableInternetFilesInPV
        Case mso_UnsafeLocationsInPV: GoSub GS¦Change_Registry_DisableUnsafeLocationsInPV
    End Select

    Call RegCloseKey(Handle): Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_DisableAttachmentsInPV:

    Call RegQueryValueEx(Handle, "DisableAttachmentsInPV", 0&, _
                         REG_DWORD, API_lpData, 4&)

    If API_lpData <> ProtectedView_State Then
        MacroSecurity_ProtectedView_Assign = RegSetValueEx(Handle, "DisableAttachmentsInPV", 0&, _
                                             REG_DWORD, ProtectedView_State, 4&) = 0&
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_DisableInternetFilesInPV:

    Call RegQueryValueEx(Handle, "DisableInternetFilesInPV", 0&, _
                         REG_DWORD, API_lpData, 4&)

    If API_lpData <> ProtectedView_State Then
        MacroSecurity_ProtectedView_Assign = RegSetValueEx(Handle, "DisableInternetFilesInPV", 0&, _
                                             REG_DWORD, ProtectedView_State, 4&) = 0&
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Change_Registry_DisableUnsafeLocationsInPV:

    Call RegQueryValueEx(Handle, "DisableUnsafeLocationsInPV", 0&, _
                         REG_DWORD, API_lpData, 4&)

    If API_lpData <> ProtectedView_State Then
        MacroSecurity_ProtectedView_Assign = RegSetValueEx(Handle, "DisableUnsafeLocationsInPV", 0&, _
                                             REG_DWORD, ProtectedView_State, 4&) = 0&
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'
    
'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


'========================================================================================================'
Public Function MacroSecurity_ProtectedView_Fetch( _
                ByVal Type_Document As MSOffice_Type_Document_Group_3, _
                ByVal Type_ProtectedView As MSOffice_Type_ProtectedView, _
                Optional ByVal Status_Code As Boolean = False _
       ) As Variant
'--------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'--------------------------------------------------------------------------------------------------------'
    
    '``````````````````````````````````````````````````````````````````````````````````````````'
    Dim Handle As LongPtr, WinAPI_Result As Long, API_lpData   As Long, API_lpSubKey As String
    Dim Section_Data As Registry_SectionData, Registry_Section As String, Sub_Handle As LongPtr
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````'
    MacroSecurity_ProtectedView_Fetch = False
    '````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document

        Case mso_Excel_3
             API_lpSubKey = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY
             Registry_Section = Registry_Section & REGISTRY_SECTION_EXCEL_SECURITY & "ProtectedView"

        Case mso_PowerPoint_3
             API_lpSubKey = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY
             Registry_Section = Registry_Section & REGISTRY_SECTION_POWERPOINT_SECURITY & "ProtectedView"

        Case mso_Word_3
             API_lpSubKey = Registry_Section & REGISTRY_SECTION_WORD_SECURITY
             Registry_Section = Registry_Section & REGISTRY_SECTION_WORD_SECURITY & "ProtectedView"

    End Select
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Section_Data = Get_Section_Data(HKEY_CURRENT_USER, Registry_Section, KEY_READ)
    If Len(Section_Data.Section_Path) = 0 Then Exit Function
    Handle = Section_Data.Handle: If Handle = 0 Then GoSub GS¦Create_DefaultKey
    '`````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_ProtectedView

        Case mso_AttachmentsInPV
            WinAPI_Result = RegQueryValueEx(Handle, "DisableAttachmentsInPV", 0&, _
                                            REG_DWORD, API_lpData, 4&)
            If WinAPI_Result = ERROR_FILE_NOT_FOUND Then GoSub GS¦Create_DefaultKey: API_lpData = 0&

        Case mso_InternetFilesInPV
            WinAPI_Result = RegQueryValueEx(Handle, "DisableInternetFilesInPV", 0&, _
                                            REG_DWORD, API_lpData, 4&)
            If WinAPI_Result = ERROR_FILE_NOT_FOUND Then GoSub GS¦Create_DefaultKey: API_lpData = 0&

        Case mso_UnsafeLocationsInPV
            WinAPI_Result = RegQueryValueEx(Handle, "DisableUnsafeLocationsInPV", 0&, _
                                            REG_DWORD, API_lpData, 4&)
            If WinAPI_Result = ERROR_FILE_NOT_FOUND Then GoSub GS¦Create_DefaultKey: API_lpData = 0&

    End Select
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````'
    If Not Status_Code Then
        Select Case API_lpData
            Case 0: MacroSecurity_ProtectedView_Fetch = "Защищенный просмотр включён"
            Case 1: MacroSecurity_ProtectedView_Fetch = "Защищенный просмотр отключён"
        End Select
    Else
        MacroSecurity_ProtectedView_Fetch = CBool(IIf(CBool(API_lpData), 0, 1))
    End If

    Call RegCloseKey(Handle)

    Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_DefaultKey:

    WinAPI_Result = RegOpenKeyEx(HKEY_CURRENT_USER, API_lpSubKey, 0&, KEY_ALL_ACCESS, Handle)
    If WinAPI_Result = 0 Then API_lpSubKey = "ProtectedView" Else Exit Function


    Call RegCreateKeyEx(Handle, API_lpSubKey, 0&, vbNullString, 0&, _
                        KEY_ALL_ACCESS, 0&, Sub_Handle, ByVal 0&)
    Call RegCloseKey(Handle): Handle = Sub_Handle

    Select Case Type_ProtectedView

        Case mso_AttachmentsInPV
                Call RegSetValueEx(Handle, "DisableAttachmentsInPV", 0&, REG_DWORD, 0&, 4&)

        Case mso_InternetFilesInPV
                Call RegSetValueEx(Handle, "DisableInternetFilesInPV", 0&, REG_DWORD, 0&, 4&)

        Case mso_UnsafeLocationsInPV
                Call RegSetValueEx(Handle, "DisableUnsafeLocationsInPV", 0&, REG_DWORD, 0&, 4&)

    End Select

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'==============================================================================================================='
Public Function MacroSecurity_TrustedLocations_Add( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                ByRef Trusted_Locations As String, _
                Optional ByRef Description As String = vbNullString, _
                Optional ByRef Allow_SubFolders As Boolean = False _
       ) As Variant
'---------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'---------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Registry_TrustedLocations    As Variant, Dict_TL_Paths As Object
    Dim WinAPI_Result    As Long, Handle As LongPtr, API_lpSubKey As String, Error_Message As String
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    MacroSecurity_TrustedLocations_Add = False
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '`````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If Get_Directory_Type(Trusted_Locations) = DirType_Invalid Then
        Error_Message = "Невозможно добавить данную директорию: " & Trusted_Locations
        Show_ErrorMessage_Immediate Error_Message, "Ошибка добавления доверенной директории!"
        Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_TrustedLocations = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Trusted Locations")
    If Registry_TrustedLocations = vbNullString Then MacroSecurity_TrustedLocations_Add = False: Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````'
    Set Dict_TL_Paths = Get_TrustedLocations_Paths(Type_Document)
    '````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If Dict_TL_Paths.Exists(Trusted_Locations) Then
        Error_Message = "Невозможно добавить данную директорию: " & Trusted_Locations
        Show_ErrorMessage_Immediate Error_Message, "Директория уже существует как доверенная!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    WinAPI_Result = RegOpenKeyEx(HKEY_CURRENT_USER, Registry_TrustedLocations, 0&, KEY_ALL_ACCESS, Handle)
    If Not WinAPI_Result = ERROR_SUCCESS Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    API_lpSubKey = Create_GUID_Microsoft()
    Call RegCreateKeyEx(Handle, API_lpSubKey, 0&, vbNullString, 0&, _
                        KEY_ALL_ACCESS, 0&, Handle, ByVal 0&)
    Call RegCloseKey(Handle)
    Call RegOpenKeyEx(HKEY_CURRENT_USER, Registry_TrustedLocations & "\" & API_lpSubKey, 0&, _
                          KEY_ALL_ACCESS, Handle)
    Call RegSetValueEx(Handle, "Path", 0&, REG_SZ, ByVal Trusted_Locations, Len(Trusted_Locations))

    If Not Description = vbNullString Then
        Call RegSetValueEx(Handle, "Description", 0&, REG_SZ, ByVal Description, Len(Description))
    End If

    If Allow_SubFolders Then Call RegSetValueEx(Handle, "AllowSubFolders", 0&, REG_DWORD, 1&, 4&)

    Call RegCloseKey(Handle): MacroSecurity_TrustedLocations_Add = True
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------------'
End Function
'==============================================================================================================='


'=================================================================================================================='
Public Function MacroSecurity_TrustedLocations_Fetch( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                Optional ByVal Show_List As Boolean = False _
       ) As Variant
'------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'------------------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data As Registry_SectionData, Dict_Sections As Object, D_Key
    Dim Matrix_Data() As Variant, Vector_SectionFullPath() As String, Inx_LB1 As Long, Inx_UB1 As Long
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, API_lpName As String, API_lpcdNameSize As Long, API_dwIndex  As Long
    Dim tmp_KeyData_1 As Registry_KeyData, tmp_KeyData_2 As Registry_KeyData, tmp_KeyData_3 As Registry_KeyData
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Table_Name As String, Table_Headers As Variant, Table_ColumnWidth As Variant, Error_Message As String
    Dim App_Name As String, Obj_SDI As Object, Obj_WB As Object, Obj_WS As Object, i As Long
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    MacroSecurity_TrustedLocations_Fetch = False
    '```````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access":     App_Name = "MS ACCESS"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel":      App_Name = "MS EXCEL"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint": App_Name = "MS POWER_POINT"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word":       App_Name = "MS WORD"
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Trusted Locations")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MacroSecurity_TrustedLocations_Fetch = False: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_Sections.keys

        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_READ)
        If Not Section_Data.Section_Count = 0 Then
            ReDim Vector_SectionFullPath(1 To Section_Data.Section_Count)
            Do
                API_lpName = String$(255, vbNullChar): API_lpcdNameSize = 255
                WinAPI_Result = RegEnumKeyEx(Section_Data.Handle, API_dwIndex, _
                                                                  API_lpName, _
                                                                  API_lpcdNameSize, _
                                                                  0&, ByVal 0&, _
                                                                      ByVal 0&, _
                                                                      ByVal 0&)
                API_dwIndex = API_dwIndex + 1

                If WinAPI_Result = ERROR_NO_MORE_ITEMS Then Exit Do
                If WinAPI_Result = ERROR_SUCCESS Then
                    Vector_SectionFullPath(API_dwIndex) = Section_Data.Section_Path & _
                                                          "\" & Left$(API_lpName, API_lpcdNameSize)
                End If
            Loop
        End If

        Call RegCloseKey(Section_Data.Handle)

    Next D_Key
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    Inx_LB1 = LBound(Vector_SectionFullPath, 1)
    Inx_UB1 = UBound(Vector_SectionFullPath, 1)

    ReDim Matrix_Data(Inx_LB1 To Inx_UB1, 1 To 4)
    '````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = Inx_LB1 To Inx_UB1

        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(Vector_SectionFullPath(i)), KEY_READ)
        tmp_KeyData_1 = Get_Key_Data(Section_Data.Handle, "Path", Section_Data.Keys_MaxLen_Value)
        tmp_KeyData_2 = Get_Key_Data(Section_Data.Handle, "AllowSubFolders", Section_Data.Keys_MaxLen_Value, True)
        tmp_KeyData_3 = Get_Key_Data(Section_Data.Handle, "Description", Section_Data.Keys_MaxLen_Value, True)

        Matrix_Data(i, 1) = "Item " & i
        Matrix_Data(i, 2) = Expand_EnvironmentVariables(CStr(tmp_KeyData_1.Key_Value))
        Matrix_Data(i, 3) = CBool(tmp_KeyData_2.Key_Value)
        Matrix_Data(i, 4) = Convert_Description(tmp_KeyData_3.Key_Value)

        Call RegCloseKey(Section_Data.Handle)

    Next i
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If Show_List Then

        Table_Name = "TRUSTED LOCATION _ " & App_Name
        Table_Headers = Array("INDEX" & vbCrLf & "Unique ID", _
                              "FOLDER LOCATION" & vbCrLf & "Path to folder (Directory)", _
                              "ROOT STATUS" & vbCrLf & "Allow SubFolders", _
                              "DESCRIPTION" & vbCrLf & "Brief description of the directory")
        Table_ColumnWidth = Array(12, 60, 25, 75)

        If Application.Name = "Microsoft Excel" Then
            If Not Create_WorkSheet_InExcel( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_TrustedLocations", "objSheet_TrustedLocations" _
            ) Then
                Error_Message = "Книга, в которую вы хотите вывести лог, защищена!" & _
                                " Создание новых листов возможно только после снятия защиты!"

                Show_ErrorMessage_Immediate Error_Message, "Структура книги защищена!"
            End If
        Else
            If Not Create_Process_ExcelApplication( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_TrustedLocations" _
            ) Then
                Error_Message = "Невозможно создать отдельный процесс приложения посредством SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Непредвиденная ошибка интерфейса SDI!"
            End If
        End If

    End If

    On Error GoTo 0: MacroSecurity_TrustedLocations_Fetch = Matrix_Data
    '`````````````````````````````````````````````````````````````````````````````````````````````'
    
'------------------------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================================='


'=================================================================================================================='
Public Function MacroSecurity_TrustedLocations_Remove( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                ByRef Trusted_Locations As String _
       ) As Boolean
'------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Registry_TrustedLocations
    Dim Dict_TL_Paths    As Object, D_Key, WinAPI_Result As Long
    Dim Dict_TL_SystemPaths As Object
    '```````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    MacroSecurity_TrustedLocations_Remove = False
    '````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '```````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````'
    If Len(Trusted_Locations) = 0 Then Exit Function
    '```````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_TrustedLocations = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Trusted Locations")
    If Registry_TrustedLocations = vbNullString Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_TL_SystemPaths = Get_Exceptions_TrustedLocations()
    Set Dict_TL_Paths = Get_TrustedLocations_Paths(Type_Document)

    For Each D_Key In Dict_TL_Paths.keys
        If Not Dict_TL_SystemPaths.Exists(D_Key) Then
            If Expand_EnvironmentVariables(CStr(D_Key)) = Expand_EnvironmentVariables(Trusted_Locations) Then
                WinAPI_Result = RegDeleteKey(HKEY_CURRENT_USER, CStr(Dict_TL_Paths(D_Key)))
                If WinAPI_Result = ERROR_SUCCESS Then MacroSecurity_TrustedLocations_Remove = True
                Exit For
            End If
        End If
    Next D_Key
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------------'
End Function
'==============================================================================================================='


'==============================================================================================================='
Public Function MacroSecurity_TrustedLocations_Reset( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2 _
       ) As Boolean
'---------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'---------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Registry_TrustedLocations, WinAPI_Result As Long
    Dim Dict_TL_SystemPaths As Object, Dict_TL_Paths As Object, D_Key
    '```````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    MacroSecurity_TrustedLocations_Reset = False
    '```````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_TrustedLocations = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Trusted Locations")
    If Registry_TrustedLocations = vbNullString Then MacroSecurity_TrustedLocations_Reset = False: Exit Function
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````'
    Set Dict_TL_SystemPaths = Get_Exceptions_TrustedLocations()
    Set Dict_TL_Paths = Get_TrustedLocations_Paths(Type_Document)
    '````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_TL_Paths.keys
        If Not Dict_TL_SystemPaths.Exists(D_Key) Then
            WinAPI_Result = RegDeleteKey(HKEY_CURRENT_USER, CStr(Dict_TL_Paths(D_Key)))
            If WinAPI_Result = ERROR_SUCCESS Then MacroSecurity_TrustedLocations_Reset = True
        End If
    Next D_Key
    '````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------------'
End Function
'==============================================================================================================='


'--------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'--------------------------------------------------------------------------------------------------------------------------'


'=========================================================================================================================='
Public Function MacroSecurity_TrustRecords_Add()
'--------------------------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox _
    "Данная функция (""MacroSecurity_TrustRecords_Add"") не включена в текущую сборку модуля по инциативе разработчика!", _
                                                                                      vbInformation, "[DarkSec_Project]"
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================================='


'=============================================================================================================='
Public Function MacroSecurity_TrustRecords_Fetch( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                Optional ByVal Show_List As Boolean = False _
       ) As Variant
'--------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'--------------------------------------------------------------------------------------------------------------'
    
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data() As Registry_SectionData
    Dim App_Name As String, Obj_SDI As Object, Obj_WB As Object, Obj_WS As Object
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_Sections As Object, D_Key, Array_Offset As Long, Inx As Long, i   As Long
    Dim Table_Name    As String, Table_Headers As Variant, Table_ColumnWidth() As Variant
    Dim Count_AllKeys As Long, TimeZone_GTM As Long, Error_Message As String
    '````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````'
    MacroSecurity_TrustRecords_Fetch = False
    '```````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '`````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access":     App_Name = "MS ACCESS"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel":      App_Name = "MS EXCEL"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint": App_Name = "MS POWER_POINT"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word":       App_Name = "MS WORD"
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "TrustRecords")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MacroSecurity_TrustRecords_Fetch = vbNullString: Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    ReDim Section_Data(1 To Dict_Sections.Count): Inx = 1

    For Each D_Key In Dict_Sections.keys
        Section_Data(Inx) = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_READ)
        Count_AllKeys = Count_AllKeys + Section_Data(Inx).Keys_Count: Inx = Inx + 1
    Next D_Key
    '``````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    If Count_AllKeys = 0 Then MacroSecurity_TrustRecords_Fetch = vbNullString: Exit Function
    ReDim Matrix_Data(1 To Count_AllKeys, 1 To 4)
    '```````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    For i = 1 To Dict_Sections.Count
        If Not Section_Data(i).Keys_Count = 0 Then
            Call Get_TrustRecords_Keys(Matrix_Data, Section_Data(i).Handle, _
                                           Section_Data(i).Keys_MaxLen_Value)
        End If

        Call RegCloseKey(Section_Data(i).Handle)
    Next
    '`````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Show_List Then

        Table_Name = "TRUST RECORDS _ " & App_Name
        Table_Headers = Array("INDEX" & vbCrLf & "Unique ID", _
                              "FILE NAME" & vbCrLf & "Documents FileName", _
                              "HEX ID" & vbCrLf & "Build File ID", _
                              "FILE LOCATION" & vbCrLf & "Path to folder (Directory)")
        Table_ColumnWidth = Array(12, 45, 50, 100)

        If Application.Name = "Microsoft Excel" Then
            If Not Create_WorkSheet_InExcel( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_TrustRecords", "objSheet_TrustRecords" _
            ) Then
                Error_Message = "Книга, в которую вы хотите вывести лог, защищена!" & _
                                " Создание новых листов возможно только после снятия защиты!"

                Show_ErrorMessage_Immediate Error_Message, "Структура книги защищена!", , True
            End If
        Else
            If Not Create_Process_ExcelApplication( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_TrustRecords" _
            ) Then
                Error_Message = "Невозможно создать отдельный процесс приложения посредством SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Непредвиденная ошибка интерфейса SDI!", , True
            End If
        End If

    End If

    On Error GoTo 0: MacroSecurity_TrustRecords_Fetch = Matrix_Data
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    
'-------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================='


'======================================================================================================='
Public Function MacroSecurity_TrustRecords_Remove( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                ByRef FileSystem_FullPath As String _
       ) As Boolean
'-------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data As Registry_SectionData
    Dim WinAPI_Result As Long, Dict_Sections As Object, D_Key, tmp_lpValueName As String
    Dim API_lpValueName As String, API_lpcdValueNameSize As Long, API_dwIndex  As Long
    '```````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````'
    MacroSecurity_TrustRecords_Remove = False
    '````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '```````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    If Get_Directory_Type(FileSystem_FullPath) = DirType_Invalid Then Exit Function

    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "TrustRecords")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MacroSecurity_TrustRecords_Remove = False: Exit Function
    '```````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_Sections.keys

        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_ALL_ACCESS)

        If Not Section_Data.Keys_Count = 0 Then
            Do
                API_lpValueName = String$(255, vbNullChar): API_lpcdValueNameSize = 255
                WinAPI_Result = RegEnumValue(Section_Data.Handle, API_dwIndex, _
                                                                  API_lpValueName, _
                                                                  API_lpcdValueNameSize, 0&, 0&, 0&, 0&)
                API_dwIndex = API_dwIndex + 1

                If WinAPI_Result = ERROR_SUCCESS Then
                    API_lpValueName = Left$(API_lpValueName, API_lpcdValueNameSize)
                    tmp_lpValueName = Expand_EnvironmentVariables(API_lpValueName)

                    If FileSystem_FullPath = tmp_lpValueName Then
                        Call RegDeleteValue(Section_Data.Handle, API_lpValueName)
                        Call RegCloseKey(Section_Data.Handle)
                        MacroSecurity_TrustRecords_Remove = True: Exit Do
                    End If
                ElseIf WinAPI_Result <> ERROR_MORE_DATA And WinAPI_Result <> ERROR_SUCCESS Then
                    Exit Do
                End If
            Loop
        End If

        Call RegCloseKey(Section_Data.Handle)

    Next D_Key
    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    
'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


'======================================================================================================='
Public Function MacroSecurity_TrustRecords_Reset( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2 _
       ) As Boolean
'-------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data As Registry_SectionData
    Dim WinAPI_Result    As Long, Dict_Sections  As Object, D_Key, tmp_lpValueName As String
    Dim API_lpValueName  As String, API_lpcdValueNameSize As Long, API_dwIndex     As Long
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````'
    MacroSecurity_TrustRecords_Reset = False
    '```````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '```````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "TrustRecords")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MacroSecurity_TrustRecords_Reset = False: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_Sections.keys

        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_ALL_ACCESS)

        If Not Section_Data.Keys_Count = 0 Then
            API_dwIndex = Section_Data.Keys_Count - 1
            Do
                API_lpValueName = String$(255, vbNullChar): API_lpcdValueNameSize = 255
                WinAPI_Result = RegEnumValue(Section_Data.Handle, API_dwIndex, _
                                                                  API_lpValueName, _
                                                                  API_lpcdValueNameSize, 0&, 0&, 0&, 0&)
                API_dwIndex = API_dwIndex - 1

                If WinAPI_Result = ERROR_SUCCESS Then
                    API_lpValueName = Left$(API_lpValueName, API_lpcdValueNameSize)
                    Call RegDeleteValue(Section_Data.Handle, API_lpValueName)
                ElseIf WinAPI_Result <> ERROR_MORE_DATA And WinAPI_Result <> ERROR_SUCCESS Then
                    Exit Do
                End If
            Loop
        End If

        Call RegCloseKey(Section_Data.Handle)

    Next D_Key
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


'-------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'-------------------------------------------------------------------------------------------------------------------'


'==================================================================================================================='
Public Function MRU_ViewableItems_Fetch_Files( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                Optional ByVal Show_List As Boolean = False _
       ) As Variant
'------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data() As Registry_SectionData
    Dim App_Name As String, Obj_SDI As Object, Obj_WB As Object, Obj_WS As Object
    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_Sections As Object, D_Key, Array_Offset As Long, Inx As Long, i As Long
    Dim Table_Name As String, Table_Headers As Variant, Table_ColumnWidth() As Variant
    Dim Count_AllKeys As Long, TimeZone_GTM As Long, Error_Message As String
    '`````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    MRU_ViewableItems_Fetch_Files = False
    '````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '`````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access":     App_Name = "MS ACCESS"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel":      App_Name = "MS EXCEL"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint": App_Name = "MS POWER_POINT"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word":       App_Name = "MS WORD"
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "File MRU")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section & "\User MRU", "File MRU")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MRU_ViewableItems_Fetch_Files = vbNullString: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````'
    ReDim Section_Data(1 To Dict_Sections.Count): Inx = 1

    For Each D_Key In Dict_Sections.keys
        Section_Data(Inx) = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_READ)
        Count_AllKeys = Count_AllKeys + Section_Data(Inx).Keys_Count: Inx = Inx + 1
    Next D_Key
    '`````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    TimeZone_GTM = Get_TimeZone_GMT: ReDim Matrix_Data(1 To Count_AllKeys, 1 To 4)
    '`````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = 1 To Dict_Sections.Count
        If Not Section_Data(i).Keys_Count = 0 Then
            Call Get_MRU_Keys(Matrix_Data, Section_Data(i).Handle, _
                                           Section_Data(i).Keys_MaxLen_Value, _
                                           TimeZone_GTM, Array_Offset)
        End If

        Array_Offset = Array_Offset + Section_Data(i).Keys_Count

        Call RegCloseKey(Section_Data(i).Handle)
    Next

    If Show_List Then

        Table_Name = "VIEWABLE FILES _ " & App_Name
        Table_Headers = Array("INDEX" & vbCrLf & "Unique ID", _
                          "FILE NAME" & vbCrLf & "Documents FileName", _
                          "FILE TIME" & vbCrLf & "File opening time", _
                          "FILE LOCATION" & vbCrLf & "Path to folder file (Document)")
        Table_ColumnWidth = Array(12, 45, 20, 100)

        If Application.Name = "Microsoft Excel" Then
            If Not Create_WorkSheet_InExcel( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_ViewableFiles", "objSheet_ViewableFiles" _
            ) Then
                Error_Message = "Книга, в которую вы хотите вывести лог, защищена!" & _
                                " Создание новых листов возможно только после снятия защиты!"

                Show_ErrorMessage_Immediate Error_Message, "Структура книги защищена!", , True
            End If
        Else
            If Not Create_Process_ExcelApplication( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_ViewableFiles" _
            ) Then
                Error_Message = "Невозможно создать отдельный процесс приложения посредством SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Непредвиденная ошибка интерфейса SDI!", , True
            End If
        End If

    End If

    On Error GoTo 0: MRU_ViewableItems_Fetch_Files = Matrix_Data
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'==================================================================================================================='
Public Function MRU_ViewableItems_Fetch_Folders( _
                ByVal Type_Document As MSOffice_Type_Document_Group_2, _
                Optional ByVal Show_List As Boolean = False _
       ) As Variant
'-------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/vbd_kit_security.197
'-------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Registry_Section As String, Section_Data() As Registry_SectionData
    Dim App_Name As String, Obj_SDI As Object, Obj_WB As Object, Obj_WS As Object
    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_Sections As Object, D_Key, Array_Offset As Long, Inx As Long, i As Long
    Dim Table_Name As String, Table_Headers As Variant, Table_ColumnWidth() As Variant
    Dim Count_AllKeys As Long, TimeZone_GTM As Long, Error_Message As String
    '`````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````'
    MRU_ViewableItems_Fetch_Folders = False
    '``````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````'
    Call Init_VBD_Kit_Security
    '`````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Registry_Section = REGISTRY_SECTION_MSO & Split(Application.Version, ".")(0) & "." _
                                            & Split(Application.Version, ".")(1)

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access":     App_Name = "MS ACCESS"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel":      App_Name = "MS EXCEL"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint": App_Name = "MS POWER_POINT"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word":       App_Name = "MS WORD"
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Place MRU")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section & "\User MRU", "Place MRU")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then MRU_ViewableItems_Fetch_Folders = vbNullString: Exit Function
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    ReDim Section_Data(1 To Dict_Sections.Count): Inx = 1

    For Each D_Key In Dict_Sections.keys
        Section_Data(Inx) = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_READ)
        Count_AllKeys = Count_AllKeys + Section_Data(Inx).Keys_Count: Inx = Inx + 1
    Next D_Key
    '``````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    TimeZone_GTM = Get_TimeZone_GMT: ReDim Matrix_Data(1 To Count_AllKeys, 1 To 4)
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = 1 To Dict_Sections.Count
        If Not Section_Data(i).Keys_Count = 0 Then
            Call Get_MRU_Keys(Matrix_Data, Section_Data(i).Handle, _
                                           Section_Data(i).Keys_MaxLen_Value, _
                                           TimeZone_GTM, Array_Offset)
        End If

        Array_Offset = Array_Offset + Section_Data(i).Keys_Count

        Call RegCloseKey(Section_Data(i).Handle)
    Next

    If Show_List Then

        Table_Name = "VIEWABLE FOLDERS _ " & App_Name
        Table_Headers = Array("INDEX" & vbCrLf & "Unique ID", _
                              "FOLDER NAME" & vbCrLf & "-", _
                              "FOLDER TIME" & vbCrLf & "Folder opening time", _
                              "FOLDER LOCATION" & vbCrLf & "Path to folder (Directory)")
        Table_ColumnWidth = Array(12, 45, 20, 100)

        If Application.Name = "Microsoft Excel" Then
            If Not Create_WorkSheet_InExcel( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_ViewableFolders", "objSheet_ViewableFolders" _
            ) Then
                Error_Message = "Книга, в которую вы хотите вывести лог, защищена!" & _
                                " Создание новых листов возможно только после снятия защиты!"

                Show_ErrorMessage_Immediate Error_Message, "Структура книги защищена!", , True
            End If
        Else
            If Not Create_Process_ExcelApplication( _
                   Matrix_Data, Table_Headers, Table_Name, _
                   Table_ColumnWidth, "db_ViewableFolders" _
            ) Then
                Error_Message = "Невозможно создать отдельный процесс приложения посредством SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Непредвиденная ошибка интерфейса SDI!", , True
            End If
        End If

    End If

    On Error GoTo 0: MRU_ViewableItems_Fetch_Folders = Matrix_Data
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


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
    Call ShellExecute(ByVal 0&, "open", Folder_ProcessedFiles, "", "", SW_SHOWNORMAL)
    '````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------'
End Sub
'===================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'==========================================================================='
Private Sub Init_VBD_Kit_Security()
'---------------------------------------------------------------------------'

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

'---------------------------------------------------------------------------'
End Sub
'==========================================================================='


'=============================================================================================='
Private Sub Init_AccessObjectModel( _
            Optional ByVal Type_Document As MSOffice_Type_Document_Group_1 = mso_Excel_1, _
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
                                                                      vbInformation, "[DarkSec_Project]"

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
                                                                      vbInformation, "[DarkSec_Project]"

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


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


'============================================================================================='
Private Function Get_Key_Data( _
                 ByRef Handle As LongPtr, _
                 ByRef Key_Name As String, _
                 ByVal Key_Data_MaxLen As Long, _
                 Optional ByVal Ignore_Errors As Boolean = False _
        ) As Registry_KeyData
'---------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, Error_Message As String, Key_Data() As Byte, i As Long
    Dim API_lpType As Long, API_lpcbData As Long, API_dwSize As Long, API_lpData As String
    Dim Len_API_lpData As Long, API_dwData As Long, API_qwData As Long
    '`````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    WinAPI_Result = RegQueryValueEx(Handle, Key_Name, 0&, API_lpType, ByVal 0&, API_lpcbData)

    If Not WinAPI_Result = ERROR_SUCCESS Then

        If Ignore_Errors Then Exit Function

        Error_Message = Get_ErrorMessage_WinAPI(WinAPI_Result)

        Show_ErrorMessage_Immediate Error_Message, _
                                    "Ошибка при вызове WinAPI", "advapi32.dll_RegQueryInfoKey"
        Exit Function

    End If
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    Get_Key_Data.Key_Name = Key_Name: Get_Key_Data.Key_Type = API_lpType
    API_dwSize = Key_Data_MaxLen: API_lpData = String(API_dwSize, vbNullChar)
    Call RegQueryValueEx(Handle, Key_Name, 0&, API_lpType, ByVal API_lpData, API_lpcbData)
    '`````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Select Case API_lpType
        Case REG_BINARY
            ReDim Key_Data(API_lpcbData - 1)
            Call RegQueryValueEx(Handle, Key_Name, 0, API_lpType, Key_Data(0), API_lpcbData)
            Get_Key_Data.Key_Value = Key_Data

        Case REG_DWORD
            CopyMemory API_dwData, ByVal API_lpData, Len(API_dwData)
            Get_Key_Data.Key_Value = API_dwData

        Case REG_QWORD
            CopyMemory API_qwData, ByVal API_lpData, Len(API_qwData)
            Get_Key_Data.Key_Value = API_qwData

        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            Get_Key_Data.Key_Value = Left$(API_lpData, API_lpcbData - 1)
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------'
End Function
'============================================================================================='


'======================================================================================================================='
Private Sub Get_MRU_Keys( _
            ByRef Matrix_Data As Variant, _
            ByVal Section_Handle As LongPtr, _
            ByVal API_dwSize As Long, _
            ByVal TimeZone_GTM As Long, _
            ByVal Array_Offset As Long _
        )
'-----------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, tmp_KeyData As Registry_KeyData, tmp_Inx As Long
    Dim tmp_Time As String, tmp_FinalTime As Date, tmp_SystemTime As WinAPI_SystemTime
    '`````````````````````````````````````````````````````````````````````````````````'
    Dim API_lpValueName As String, API_lpcdValueNameSize As Long
    Dim API_dwIndex As Long, API_lpFileTime As WinAPI_FileTime
    '`````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Do

        API_lpValueName = String$(255, vbNullChar): API_lpcdValueNameSize = 255
        WinAPI_Result = RegEnumValue(Section_Handle, API_dwIndex, _
                                                     API_lpValueName, _
                                                     API_lpcdValueNameSize, 0&, 0&, 0&, 0&)
        API_dwIndex = API_dwIndex + 1

        If WinAPI_Result = ERROR_SUCCESS Then

            API_lpValueName = Left$(API_lpValueName, API_lpcdValueNameSize)

            tmp_KeyData = Get_Key_Data(Section_Handle, API_lpValueName, API_dwSize)

            tmp_Inx = InStrRev(tmp_KeyData.Key_Value, "\")
            tmp_Time = Mid$(tmp_KeyData.Key_Value, InStr(1, tmp_KeyData.Key_Value, "][") + 3, 16)

            On Error Resume Next

            API_lpFileTime.tm_LowDateTime = CDec("&H" & Right$(tmp_Time, 8))
            API_lpFileTime.tm_HighDateTime = CDec("&H" & Left$(tmp_Time, 8))

            If Not Err.Number = 0 Then

                On Error GoTo 0
                tmp_Inx = InStrRev(tmp_KeyData.Key_Value, "\", Len(tmp_KeyData.Key_Value) - 1)
                Matrix_Data(API_dwIndex + Array_Offset, 1) = "-"
                Matrix_Data(API_dwIndex + Array_Offset, 2) = "Undefined_Key"
                Matrix_Data(API_dwIndex + Array_Offset, 3) = DateSerial(2015, 1, 1) + TimeSerial(12, 0, 0)
                Matrix_Data(API_dwIndex + Array_Offset, 4) = tmp_KeyData.Key_Value

            Else

                tmp_SystemTime = Convert_FileTimeToSystemTime(API_lpFileTime)

                tmp_FinalTime = DateSerial(tmp_SystemTime.tm_Year, _
                                           tmp_SystemTime.tm_Month, _
                                           tmp_SystemTime.tm_Day)

                tmp_FinalTime = tmp_FinalTime + TimeSerial(tmp_SystemTime.tm_Hour + TimeZone_GTM, _
                                                           tmp_SystemTime.tm_Minute, _
                                                           tmp_SystemTime.tm_Second)

                tmp_Inx = InStrRev(tmp_KeyData.Key_Value, "\", Len(tmp_KeyData.Key_Value) - 1)

                Matrix_Data(API_dwIndex + Array_Offset, 1) = tmp_KeyData.Key_Name
                Matrix_Data(API_dwIndex + Array_Offset, 2) = Mid$(tmp_KeyData.Key_Value, 1 + tmp_Inx, _
                                                              Len(tmp_KeyData.Key_Value) - tmp_Inx)
                Matrix_Data(API_dwIndex + Array_Offset, 3) = tmp_FinalTime
                Matrix_Data(API_dwIndex + Array_Offset, 4) = Mid$(tmp_KeyData.Key_Value, InStr(1, _
                                                                  tmp_KeyData.Key_Value, "*") + 1, tmp_Inx)

                If Right$(Matrix_Data(API_dwIndex + Array_Offset, 2), 1) = "\" Then
                    Matrix_Data(API_dwIndex + Array_Offset, 2) = Left$(Matrix_Data(API_dwIndex + Array_Offset, 2), _
                                                                   Len(Matrix_Data(API_dwIndex + Array_Offset, 2)) - 1)
                End If

            End If

        ElseIf WinAPI_Result <> ERROR_MORE_DATA And WinAPI_Result <> ERROR_SUCCESS Then

            Exit Do

        End If

    Loop
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------'
End Sub
'======================================================================================================================='


'======================================================================================================================='
Private Sub Get_TrustRecords_Keys( _
            ByRef Matrix_Data As Variant, _
            ByVal Section_Handle As LongPtr, _
            ByVal API_dwSize As Long _
        )
'-----------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, tmp_KeyData As Registry_KeyData, tmp_Inx As Long
    Dim tmp_Time As String, tmp_FinalTime As Date, tmp_SystemTime As WinAPI_SystemTime
    '`````````````````````````````````````````````````````````````````````````````````'
    Dim API_lpValueName As String, API_lpcdValueNameSize As Long
    Dim API_dwIndex As Long, API_lpFileTime As WinAPI_FileTime
    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Vector_Bytes() As Byte, Inx_Offset As Long
    '`````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    Do
        API_lpValueName = String$(255, vbNullChar): API_lpcdValueNameSize = 255
        WinAPI_Result = RegEnumValue(Section_Handle, API_dwIndex, _
                                                     API_lpValueName, _
                                                     API_lpcdValueNameSize, 0&, 0&, 0&, 0&)
        API_dwIndex = API_dwIndex + 1

        If WinAPI_Result = ERROR_SUCCESS Then
            API_lpValueName = Left$(API_lpValueName, API_lpcdValueNameSize)

            tmp_KeyData = Get_Key_Data(Section_Handle, API_lpValueName, API_dwSize)
            tmp_KeyData.Key_Name = Expand_EnvironmentVariables(CStr(tmp_KeyData.Key_Name))


            tmp_Inx = InStrRev(tmp_KeyData.Key_Name, "\")
            tmp_Time = Mid$(tmp_KeyData.Key_Name, InStr(1, tmp_KeyData.Key_Name, "][") + 3, 16)

            If Not tmp_Inx = 0 Then

                Vector_Bytes = tmp_KeyData.Key_Value

                Matrix_Data(API_dwIndex - Inx_Offset, 1) = "Item " & API_dwIndex
                Matrix_Data(API_dwIndex - Inx_Offset, 2) = Mid$(tmp_KeyData.Key_Name, 1 + tmp_Inx, _
                                                            Len(tmp_KeyData.Key_Name) - tmp_Inx)
                Matrix_Data(API_dwIndex - Inx_Offset, 3) = Convert_ByteArrayToHex(Vector_Bytes)
                Matrix_Data(API_dwIndex - Inx_Offset, 4) = tmp_KeyData.Key_Name

            Else

                Inx_Offset = Inx_Offset + 1

            End If

        ElseIf WinAPI_Result <> ERROR_MORE_DATA And WinAPI_Result <> ERROR_SUCCESS Then
            Exit Do
        End If

    Loop
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------'
End Sub
'======================================================================================================================='


'======================================================================================================================='
Private Function Get_TrustedLocations_Paths( _
                 ByVal Type_Document As MSOffice_Type_Document_Group_2 _
        ) As Object
'-----------------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_TrustedLocations_Paths As Object, API_dwIndex As Long, API_lpName As String, API_lpcdNameSize As Long
    Dim Registry_Section As String, Section_Data As Registry_SectionData, tmp_KeyData As Registry_KeyData
    Dim Dict_Sections    As Object, D_Key, WinAPI_Result, Inx_LB1 As Long, Inx_UB1 As Long, i As Long
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    Registry_Section = "SOFTWARE\Microsoft\Office\" & Application.Version

    Select Case Type_Document
        Case mso_Access_2:     Registry_Section = Registry_Section & "\Access"
        Case mso_Excel_2:      Registry_Section = Registry_Section & "\Excel"
        Case mso_PowerPoint_2: Registry_Section = Registry_Section & "\PowerPoint"
        Case mso_Word_2:       Registry_Section = Registry_Section & "\Word"
    End Select
    '``````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````'
    Set Dict_Sections = CreateObject("Scripting.Dictionary")

    D_Key = Registry_Section_Find(HKEY_CURRENT_USER, Registry_Section, "Trusted Locations")
    If Not D_Key = vbNullString Then Dict_Sections(D_Key) = vbNullString

    If Dict_Sections.Count = 0 Then Get_TrustedLocations_Paths = False: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_Sections.keys
        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(D_Key), KEY_READ)

        If Not Section_Data.Section_Count = 0 Then

            ReDim Vector_SectionFullPath(1 To Section_Data.Section_Count)

            Do
                API_lpName = String$(255, vbNullChar): API_lpcdNameSize = 255
                WinAPI_Result = RegEnumKeyEx(Section_Data.Handle, API_dwIndex, _
                                                                  API_lpName, _
                                                                  API_lpcdNameSize, _
                                                                  0&, ByVal 0&, _
                                                                      ByVal 0&, _
                                                                      ByVal 0&)
                API_dwIndex = API_dwIndex + 1

                If WinAPI_Result = ERROR_NO_MORE_ITEMS Then Exit Do
                If WinAPI_Result = ERROR_SUCCESS Then
                    Vector_SectionFullPath(API_dwIndex) = Section_Data.Section_Path & _
                                                          "\" & Left$(API_lpName, API_lpcdNameSize)
                End If
            Loop
        End If

        Call RegCloseKey(Section_Data.Handle)

    Next D_Key
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    Inx_LB1 = LBound(Vector_SectionFullPath, 1)
    Inx_UB1 = UBound(Vector_SectionFullPath, 1)

    Set Dict_TrustedLocations_Paths = CreateObject("Scripting.Dictionary")
    '````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = Inx_LB1 To Inx_UB1
        Section_Data = Get_Section_Data(HKEY_CURRENT_USER, CStr(Vector_SectionFullPath(i)), KEY_READ)
        tmp_KeyData = Get_Key_Data(Section_Data.Handle, "Path", Section_Data.Keys_MaxLen_Value)
        Dict_TrustedLocations_Paths(tmp_KeyData.Key_Value) = Section_Data.Section_Path
        Dict_TrustedLocations_Paths(Expand_EnvironmentVariables( _
                               CStr(tmp_KeyData.Key_Value))) = Section_Data.Section_Path
        Call RegCloseKey(Section_Data.Handle)
    Next i
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Set Get_TrustedLocations_Paths = Dict_TrustedLocations_Paths
    '```````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================================='


'========================================================================================================'
Private Function Registry_Section_Find( _
                 ByVal Registry_HKey As LongPtr, _
                 ByRef Registry_Section As String, _
                 ByRef Section_Find As String _
        ) As String
'--------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, Handle As LongPtr, API_dwIndex As Long
    Dim Name_Section As String, Length_Name_Section As Long, Section_FullPath As String
    '``````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    WinAPI_Result = RegOpenKeyEx(Registry_HKey, Registry_Section, 0&, KEY_READ, Handle)
    Registry_Section_Find = vbNullString: If Not WinAPI_Result = ERROR_SUCCESS Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    Do
        Length_Name_Section = 255: Name_Section = Space$(Length_Name_Section)
        WinAPI_Result = RegEnumKeyEx(Handle, API_dwIndex, Name_Section, _
                                                          Length_Name_Section, _
                                                          0&, ByVal 0&, _
                                                              ByVal 0&, _
                                                              ByVal 0&)
        If WinAPI_Result = ERROR_SUCCESS Then
            Name_Section = Left$(Name_Section, Length_Name_Section)
            Section_FullPath = Registry_Section & "\" & Name_Section

            If Name_Section = Section_Find Then
                Registry_Section_Find = Section_FullPath: Exit Function
            End If

            Registry_Section_Find = Registry_Section_Find(Registry_HKey, Section_FullPath, Section_Find)
            API_dwIndex = API_dwIndex + 1: If Not Registry_Section_Find = vbNullString Then Exit Function
        End If

    Loop While WinAPI_Result = ERROR_SUCCESS

    Call RegCloseKey(Handle)
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'=================================================================='
Private Function Get_CodeRow_InModule( _
                 ByRef Code_Module As Object, _
                 ByRef Inx_StartLine As Long, _
                 ByRef Text_ToFind As String, _
                 Optional ByVal Reverse_Find As Boolean = False _
        ) As Long
'-----------------------------------------------------------------'

    '`````````````````````````````````````'
    Dim Code_LineText As String, i As Long
    '`````````````````````````````````````'

    '```````````````````````````````````````````````````````'
    If Not Reverse_Find Then
        For i = Inx_StartLine To Code_Module.CountOfLines
            Code_LineText = Code_Module.Lines(i, 1)
            If InStr(1, Code_LineText, Text_ToFind) > 0 Then
                Get_CodeRow_InModule = i + 2: Exit Function
            End If
        Next
    Else
        For i = Inx_StartLine To 1 Step -1
            Code_LineText = Code_Module.Lines(i, 1)
            If InStr(1, Code_LineText, Text_ToFind) > 0 Then
                Get_CodeRow_InModule = i - 1: Exit Function
            End If
        Next
    End If
    '```````````````````````````````````````````````````````'

'-----------------------------------------------------------------'
End Function
'=================================================================='


'========================================================================================='
Private Function Get_Function_InModule( _
                 ByRef Std_Module As Object, _
                 ByRef Code_Module As Object, _
                 ByRef Name_Function As String _
        ) As String
'-----------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````'
    Dim Vector_CodeFunction() As String, Module_TotalLines As Long, i As Long
    Dim Inx As Long, Inx_Line As Long, Inx_Start As Long, Inx_End As Long
    Dim Module_CurrentProcedure As String
    '`````````````````````````````````````````````````````````````````````````'
    Const Std_Module_ID As Long = 1
    '`````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    If Not Std_Module.Type = Std_Module_ID Then
        Get_Function_InModule = vbNullString: Exit Function
    End If
    '```````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    Module_TotalLines = Code_Module.CountOfLines: Inx_Line = 1
    '`````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````'
    Do Until Inx_Line > Module_TotalLines
        Module_CurrentProcedure = Code_Module.ProcOfLine(Inx_Line, 0&)
        If Module_CurrentProcedure <> "" Then
            If Module_CurrentProcedure = Name_Function Then
                Inx_Start = Get_CodeRow_InModule(Code_Module, Inx_Line, Name_Function) - 3
                Inx_End = Get_CodeRow_InModule(Code_Module, Inx_Start, "End Function")
                Exit Do
            Else
                Inx_Line = Inx_Line + _
                           Code_Module.ProcCountLines(Module_CurrentProcedure, 0&)
            End If
        Else
            Inx_Line = Inx_Line + 1
        End If
    Loop
    '`````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    ReDim Vector_CodeFunction(1 To Inx_End - Inx_Start)

    For i = Inx_Start To Inx_End - 1
        Inx = Inx + 1: Vector_CodeFunction(Inx) = Code_Module.Lines(i, 1)
    Next i
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````'
    Get_Function_InModule = Join(Vector_CodeFunction, vbCrLf)
    '```````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'==================================================================================='
Private Function Get_WinAPI_InModule( _
                 ByRef Code_Module As Object, _
                 ByRef Declare_WinAPI As String _
        ) As String
'-----------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Inx As Long, Inx_Start As Long, Inx_End As Long
    '``````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    #If Has_PtrSafe Then
        Inx = Get_CodeRow_InModule(Code_Module, 1&, Declare_WinAPI)
        Inx_Start = Get_CodeRow_InModule(Code_Module, Inx, "Declare PtrSafe", True)
    #Else
        Inx = Get_CodeRow_InModule(Code_Module, 1&, Declare_WinAPI)
        Inx = Get_CodeRow_InModule(Code_Module, Inx, Declare_WinAPI)
        Inx_Start = Get_CodeRow_InModule(Code_Module, Inx, "Declare", True)
    #End If

    If Declare_WinAPI Like "Function" Then
        Inx_End = Get_CodeRow_InModule(Code_Module, Inx, ") As")
    Else
        Inx_End = Get_CodeRow_InModule(Code_Module, Inx, ")")
    End If
    '```````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Get_WinAPI_InModule = Get_Text_InModule(Code_Module, Inx_Start, Inx_End)
    Get_WinAPI_InModule = Remove_LeadingChars( _
                          Get_WinAPI_InModule, Chr$(10) & Chr$(13) & Chr$(32))
    '`````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------'
End Function
'==================================================================================='


'========================================================================'
Private Function Get_Text_InModule( _
                 ByRef Code_Module As Object, _
                 ByRef Inx_StartLine As Long, _
                 ByRef Inx_EndLine As Long _
        ) As String
'------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Vector_CodeFunction() As String, i As Long, Inx As Long
    '``````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    ReDim Vector_CodeFunction(1 To Inx_EndLine - Inx_StartLine)

    For i = Inx_StartLine To Inx_EndLine - 1
        Inx = Inx + 1: Vector_CodeFunction(Inx) = Code_Module.Lines(i, 1)
    Next i
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````'
    Get_Text_InModule = Join(Vector_CodeFunction, vbCrLf)
    '```````````````````````````````````````````````````````'

'------------------------------------------------------------------------'
End Function
'========================================================================'


'==============================================================================================='
Private Function Remove_LeadingChars( _
                 ByRef Source_String As String, _
                 Optional LTrim_Chars As String = vbNullString, _
                 Optional RTrim_Chars As String = vbNullString _
        ) As String
'-----------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````'
    Dim Length_SourceString As Long, Inx_LT As Long, Inx_RT As Long
    '``````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````'
    Length_SourceString = Len(Source_String)
    If Length_SourceString = 0 Then
        Remove_LeadingChars = Source_String: Exit Function
    Else
        Inx_LT = 1: Inx_RT = Length_SourceString
    End If
    '`````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    If Not LTrim_Chars = vbNullString Then
        Do Until InStr(LTrim_Chars, Mid$(Source_String, Inx_LT, 1)) = 0
            Inx_LT = Inx_LT + 1: If Inx_LT > Inx_RT Then Remove_LeadingChars = "": Exit Function
        Loop
    End If


    If Not RTrim_Chars = vbNullString Then
        Do Until InStr(RTrim_Chars, Mid$(Source_String, Inx_RT, 1)) = 0
            Inx_RT = Inx_RT - 1: If Inx_LT > Inx_RT Then Remove_LeadingChars = "": Exit Function
        Loop
    End If
    '```````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    Remove_LeadingChars = Mid$(Source_String, Inx_LT, Inx_RT - Inx_LT + 1)
    '`````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------'
End Function
'==============================================================================================='


'===================================================='
Private Function Is_VBProjectProtectionRequired( _
                 ByRef Obj_WB As Object _
        ) As Boolean
'----------------------------------------------------'

    '``````````````````````````'
    Dim Obj_VBProject As Object
    Dim Obj_VBComp    As Object
    Dim Obj_WS        As Object
    Dim Flag_Code    As Boolean
    '``````````````````````````'

    '`````````````````````````````````````'
    Is_VBProjectProtectionRequired = False
    '`````````````````````````````````````'

    '`````````````````````````````````````````'
    On Error Resume Next
    Set Obj_VBProject = Obj_WB.VBProject
    On Error GoTo 0
    '`````````````````````````````````````````'

    '`````````````````````````````````````````'
    If Obj_VBProject Is Nothing Then
        Is_VBProjectProtectionRequired = True
        Exit Function
    End If
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````'
    On Error Resume Next
    For Each Obj_VBComp In Obj_VBProject.VBComponents
        Select Case Obj_VBComp.Type
            Case 1, 2, 3
                Is_VBProjectProtectionRequired = True
                Exit Function
        End Select
    Next Obj_VBComp
    On Error GoTo 0
    '````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    Flag_Code = Is_ThereCode(ThisWorkbook)
    If Flag_Code Then
        Is_VBProjectProtectionRequired = True
        Exit Function
    End If
    '`````````````````````````````````````````'

    '````````````````````````````````````````````'
    For Each Obj_WS In Obj_WB.Sheets
        Flag_Code = Is_ThereCode(Obj_WS)
        If Flag_Code Then
            Is_VBProjectProtectionRequired = True
            Exit Function
        End If
    Next Obj_WS
    '````````````````````````````````````````````'

'----------------------------------------------------'
End Function
'===================================================='


'============================================================='
Private Function Is_ThereCode( _
                 ByRef Obj_WS As Object _
        ) As Boolean
'-------------------------------------------------------------'

    '````````````````````````'
    Dim Obj_CodeMod As Object
    Dim Code_Text   As String
    '````````````````````````'

    '``````````````````````````````````'
    Is_ThereCode = False

    On Error Resume Next
    Set Obj_CodeMod = Obj_WS.CodeModule
    On Error GoTo 0
    '``````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    If Obj_CodeMod Is Nothing Then Exit Function

    On Error Resume Next
    Code_Text = Obj_CodeMod.Lines(1, Obj_CodeMod.CountOfLines)
    On Error GoTo 0

    If Code_Text = "" Then Exit Function
    '`````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    If InStr(1, Code_Text, "Sub", vbTextCompare) > 0 Or _
       InStr(1, Code_Text, "Function", vbTextCompare) > 0 Then
       Is_ThereCode = True
    End If
    '`````````````````````````````````````````````````````````'

'-------------------------------------------------------------'
End Function
'============================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


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


'======================================================================================================'
Private Function Show_InfoMessage_UnixNotSupported()
'------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox "Данный программный модуль не поддерживает работу в UNIX-подобных системах (MacOS)! " & _
       String$(2, vbNewLine) & "Реализована поддержка только в Windows_NT системах!" & vbNewLine & _
      "Дальнейшая работа функции будет прервана!", vbInformation, "[DarkSec_Project]"
    '```````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'================================================================================================'
Private Function Get_TimeZone_GMT() As Long
'------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````'
    Dim ActiveTimeBias As Long, GMT_Hours As Long
    '````````````````````````````````````````````'
    Const Registry_Key_ActiveTimeBias As String = _
                           "HKEY_LOCAL_MACHINE" & _
                           "\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    ActiveTimeBias = CreateObject("WScript.Shell").RegRead(Registry_Key_ActiveTimeBias)
    '`````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    If Not Err.Number = 0 Then ActiveTimeBias = 0

    GMT_Hours = ActiveTimeBias \ 60

    If ActiveTimeBias > 0 Then
        Get_TimeZone_GMT = GMT_Hours
    Else
        Get_TimeZone_GMT = -(GMT_Hours)
    End If

    On Error GoTo 0
    '````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------'
End Function
'================================================================================================'


'============================================================================================='
Private Function Get_Exceptions_TrustedLocations() As Object
'---------------------------------------------------------------------------------------------'

    '``````````````````````````````````````'
    Dim Dict_TL As Object, D_Key As Variant
    '``````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    Set Dict_TL = CreateObject("Scripting.Dictionary")

    Dict_TL("%APPDATA%\Microsoft\Templates") = vbNullString
    Dict_TL("C:\Program Files\Microsoft Office\Root\Templates\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Templates\") = vbNullString
    Dict_TL("%APPDATA%\Microsoft\Word\Startup") = vbNullString
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Dict_TL("C:\Program Files\Microsoft Office\Root\Office16\Library\") = vbNullString
    Dict_TL("C:\Program Files\Microsoft Office\Root\Office16\STARTUP\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Office16\Library\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Office16\STARTUP\") = vbNullString
    Dict_TL("%APPDATA%\Microsoft\Excel\XLSTART") = vbNullString
    Dict_TL("C:\Program Files\Microsoft Office\Root\Office16\XLSTART\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Office16\XLSTART\") = vbNullString
    '```````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Dict_TL("%APPDATA%\Microsoft\Addins") = vbNullString
    Dict_TL("C:\Program Files\Microsoft Office\Root\Document Themes 16\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\") = vbNullString
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````'
    Dict_TL("C:\Program Files\Microsoft Office\Root\Office16\ACCWIZ\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Root\Office16\ACCWIZ\") = vbNullString

    Dict_TL("C:\Program Files\Microsoft Office\Office16\ACCWIZ\") = vbNullString
    Dict_TL("C:\Program Files (x86)\Microsoft Office\Office16\ACCWIZ\") = vbNullString

    For Each D_Key In Dict_TL.keys
        Dict_TL(Expand_EnvironmentVariables(CStr(D_Key))) = vbNullString
    Next D_Key

    Set Get_Exceptions_TrustedLocations = Dict_TL
    '``````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------'
End Function
'============================================================================================='


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


'================================================================================='
Private Function Get_List_Files_ToProcess( _
                 ByVal Allow_MultiSelect As Boolean, _
                 ParamArray Templates_ID() As Variant _
        ) As Variant
'---------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````'
    Dim GUI_FileDialog As Object, Has_Filter_AllFiles As Boolean
    Dim Inx_LB1 As Long, Inx_UB1 As Long, i As Long
    Dim C_Item  As Variant, Inx  As Long
    '````````````````````````````````````````````````````````````'
    Const mso_FileDialogFilePicker = 3
    '````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Inx_LB1 = LBound(Templates_ID, 1)
    Inx_UB1 = UBound(Templates_ID, 1)

    For i = Inx_LB1 To Inx_UB1
        Select Case VarType(Templates_ID(i))
            Case vbByte, vbInteger, vbLong
            Case Else: Exit Function
        End Select
    Next i
    '`````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    Set GUI_FileDialog = Application.FileDialog(mso_FileDialogFilePicker)

    With GUI_FileDialog
        .Title = "Выберите файл(-ы) для обработки"
        .Filters.Clear

        If Allow_MultiSelect Then
            .AllowMultiSelect = True
        Else
            .AllowMultiSelect = False
        End If
    End With

    Has_Filter_AllFiles = True: Get_List_Files_ToProcess = False
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````'
    ReDim Vector_FilesType(Inx_LB1 To Inx_UB1), Vector_FilesFullPath(1 To 1)

    For i = Inx_LB1 To Inx_UB1

        Select Case Templates_ID(i)

            Case TID_FileSystem_AllFiles
                Has_Filter_AllFiles = False
                GoSub GS¦Add_FileSystem_AllFiles: Exit For

            Case TID_FileSystem_BIN
                Vector_FilesType(i) = "*.bin"
                GoSub GS¦Add_FileSystem_BIN

            Case TID_FileSystem_DLL
                Vector_FilesType(i) = "*.dll"
                GoSub GS¦Add_FileSystem_DLL

            Case TID_FileSystem_EXE
                Vector_FilesType(i) = "*.exe"
                GoSub GS¦Add_FileSystem_EXE

            Case TID_FileSystem_INI
                Vector_FilesType(i) = "*.ini"
                GoSub GS¦Add_FileSystem_INI

            Case TID_FileSystem_TXT
                Vector_FilesType(i) = "*.txt; *.log"
                GoSub GS¦Add_FileSystem_TXT

            Case TID_FileSystem_Access
                Vector_FilesType(i) = "*.accdb; *.mdb; *.accdt"
                GoSub GS¦Add_MSO_Access

            Case TID_FileSystem_Excel
                Vector_FilesType(i) = "*.xlsx; *.xls; *.xlsm; " & _
                                      "*.xlsb; *.xla; *.xlam; " & _
                                      "*.xltx; *.xlt; *.xltm"
                GoSub GS¦Add_MSO_Excel

            Case TID_FileSystem_PowerPoint
                Vector_FilesType(i) = "*.ppt; *.pptx; *.pptm; " & _
                                      "*.pot; *.potx; *.potm; " & _
                                      "*.ppa; *.ppam; " & _
                                      "*.pps; *.ppsx; *.ppsm"

                GoSub GS¦Add_MSO_PowerPoint

            Case TID_FileSystem_Word
                Vector_FilesType(i) = "*.doc; *.docx; *.docm; " & _
                                      "*.dot; *.dotx; *.dotm"
                GoSub GS¦Add_MSO_Word

        End Select
    Next i
    '``````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    If UBound(Templates_ID, 1) > 0 And Has_Filter_AllFiles Then
        GUI_FileDialog.Filters.Add "Все поддерживаемые файлы", _
                                    Join(Vector_FilesType, "; ")
    End If
    '```````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If GUI_FileDialog.Show = -1 Then
        For Each C_Item In GUI_FileDialog.SelectedItems
            Inx = Inx + 1:  ReDim Preserve Vector_FilesFullPath(1 To Inx)
            Vector_FilesFullPath(Inx) = GUI_FileDialog.SelectedItems(Inx)
        Next C_Item
    End If

    Set GUI_FileDialog = Nothing
    '```````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````'
    Get_List_Files_ToProcess = Vector_FilesFullPath

    Exit Function
    '``````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_AllFiles:

    GUI_FileDialog.Filters.Clear

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_BIN:

    GUI_FileDialog.Filters.Add "Двоичные файлы", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_DLL:

    GUI_FileDialog.Filters.Add "Библиотеки DLL", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_EXE:

    GUI_FileDialog.Filters.Add "Исполняемые файлы", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_INI:

    GUI_FileDialog.Filters.Add "Файлы конфигурации", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_FileSystem_TXT:

    GUI_FileDialog.Filters.Add "Текстовые файлы", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_MSO_Access:

    GUI_FileDialog.Filters.Add "Файлы Access", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_MSO_Excel:

    GUI_FileDialog.Filters.Add "Файлы Excel", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_MSO_PowerPoint:

    GUI_FileDialog.Filters.Add "Файлы PowerPoint", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦Add_MSO_Word:

    GUI_FileDialog.Filters.Add "Файлы Word", Vector_FilesType(i)

    Return
'```````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------'
End Function
'================================================================================='


'==============================================================================='
Private Function Expand_EnvironmentVariables( _
                 ByRef Env_Variable As String _
        ) As String
'-------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````'
    With CreateObject("WScript.Shell")
        Expand_EnvironmentVariables = .ExpandEnvironmentStrings(Env_Variable)
    End With
    '````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    Expand_EnvironmentVariables = Convert_URLDecode(Expand_EnvironmentVariables)
    Expand_EnvironmentVariables = Replace(Expand_EnvironmentVariables, "/", "\")
    '```````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


'=================================================================================================='
Private Function Convert_Description( _
                 ByRef Description_ID As Variant _
        ) As String
'--------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Description_ID
        Case 0:  Convert_Description = "Расположение Word по умолчанию: шаблоны пользователя"
        Case 1:  Convert_Description = "Расположение Word по умолчанию: шаблоны приложений"
        Case 2:  Convert_Description = "Расположение Word по умолчанию: автозагрузка"
        Case 3:  Convert_Description = "Расположение Excel по умолчанию: автозагрузка Excel"
        Case 4:  Convert_Description = "Расположение Excel по умолчанию: автозагрузка пользователя"
        Case 5:  Convert_Description = "Расположение Excel по умолчанию: шаблоны пользователя"
        Case 6:  Convert_Description = "Расположение Excel по умолчанию: шаблоны приложений"
        Case 7:  Convert_Description = "Расположение Excel по умолчанию: автозагрузка Office"
        Case 8:  Convert_Description = "Расположение PowerPoint по умолчанию: шаблоны"
        Case 9:  Convert_Description = "Расположение PowerPoint по умолчанию: шаблоны приложений"
        Case 10: Convert_Description = "Расположение PowerPoint по умолчанию: надстройки"
        Case 11: Convert_Description = "Расположение PowerPoint по умолчанию: темы приложений"
        Case 12: Convert_Description = "Расположение Excel по умолчанию: надстройки"
        Case "Access default location: Wizard Databases":
                 Convert_Description = "Расположение Access по умолчанию: Базы данных"
        Case Else:
                 Convert_Description = Description_ID
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'========================================================================='
Private Function Convert_FileTimeToSystemTime( _
                 ByRef API_lpFileTime As WinAPI_FileTime _
        ) As WinAPI_SystemTime
'-------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````'
    Dim WinAPI_Result As Long, API_lpSystemTime As WinAPI_SystemTime
    '```````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    WinAPI_Result = FileTimeToSystemTime(API_lpFileTime, API_lpSystemTime)
    WinAPI_Result = FileTimeToSystemTime(API_lpFileTime, API_lpSystemTime)

    Convert_FileTimeToSystemTime = API_lpSystemTime
    '``````````````````````````````````````````````'

'-------------------------------------------------------------------------'
End Function
'========================================================================='


'========================================================================='
Private Function Convert_ByteArrayToHex( _
                 Vector_Bytes() As Byte _
        ) As String
'-------------------------------------------------------------------------'

    '``````````````````````````````````````'
    Dim Hex_String As String, i  As Integer
    Dim Inx_LB1 As Long, Inx_UB1 As Long
    '``````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Inx_LB1 = LBound(Vector_Bytes, 1)
    Inx_UB1 = UBound(Vector_Bytes, 1)

    For i = Inx_LB1 To Inx_UB1
        Hex_String = Hex_String & Right("0" & Hex(Vector_Bytes(i)), 2)
    Next i

    Convert_ByteArrayToHex = Hex_String
    '`````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------'
End Function
'========================================================================='


'=========================================================================================='
Private Function Create_GUID_Microsoft() As String
'------------------------------------------------------------------------------------------'

    '`````````````````````````````'
    Dim GUID As WinAPI_GUID
    Dim RetValue As Variant
    '`````````````````````````````'
    Const GUID_Length As Long = 39
    '`````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    Call CoCreateGuid(GUID)
    Create_GUID_Microsoft = String$(GUID_Length, vbNullChar)
    RetValue = StringFromGUID2(GUID, StrPtr(Create_GUID_Microsoft), GUID_Length)
    '``````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````'
    If RetValue = GUID_Length Then Create_GUID_Microsoft = Left$(Create_GUID_Microsoft, 38)
    '``````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------'
End Function
'=========================================================================================='


'======================================================================='
Private Function Convert_ExcelExtension( _
                 ByRef Obj_Workbook As Object, _
                 ByVal Path_ToSave As String _
        ) As String
'-----------------------------------------------------------------------'

    '```````````````````````````````````````````````````'
    Dim File_Extension As String, Target_Format As Long
    '``````````````````````````````````````````````````````````'
    Const xl_WorkbookNormal              As Long = -4143&
    Const xl_OpenXMLWorkbookMacroEnabled As Long = 52&
    '``````````````````````````````````````````````````````````'
    Const xl_AddIn As Long = 18&, xl_OpenXMLAddIn As Long = 55&
    '``````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    Convert_ExcelExtension = vbNullString
    '````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    File_Extension = LCase$(Right$(Obj_Workbook.Name, 4))

    Select Case File_Extension

        Case ".xls"
             Target_Format = xl_OpenXMLWorkbookMacroEnabled
             Path_ToSave = Replace_Extension(Path_ToSave, ".xlsm")

        Case "xlsm"
             Target_Format = xl_WorkbookNormal
             Path_ToSave = Replace_Extension(Path_ToSave, ".xls")

        Case ".xla"
             Target_Format = xl_OpenXMLAddIn
             Path_ToSave = Replace_Extension(Path_ToSave, ".xlam")

        Case "xlam"
             Target_Format = xl_AddIn
             Path_ToSave = Replace_Extension(Path_ToSave, ".xla")

        Case Else: Exit Function

    End Select
    '`````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    Obj_Workbook.SaveAs FileName:=Path_ToSave, FileFormat:=Target_Format
    '````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    If Not CBool(Err.Number) Then Convert_ExcelExtension = Path_ToSave
    On Error GoTo 0
    '````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------'
End Function
'========================================================================'


'========================================================================'
Private Function Convert_WordExtension( _
                ByRef Obj_Document As Object, _
                ByVal Path_ToSave As String _
        ) As String
'------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim File_Extension As String, Target_Format As Long
    '``````````````````````````````````````````````````'
    Const wd_Document97 As Long = 0&
    Const wd_Template97 As Long = 1&
    Const wd_XMLDocumentMacroEnabled As Long = 13&
    Const wd_XMLTemplateMacroEnabled As Long = 15&
    '``````````````````````````````````````````````````'

    '```````````````````````````````````'
    Convert_WordExtension = vbNullString
    '```````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    File_Extension = LCase$(Right$(Obj_Document.Name, 4))

    Select Case File_Extension

        Case ".doc"
            Target_Format = wd_XMLDocumentMacroEnabled
            Path_ToSave = Replace_Extension(Path_ToSave, ".docm")

        Case "docm"
            Target_Format = wd_Document97
            Path_ToSave = Replace_Extension(Path_ToSave, ".doc")

        Case ".dot"
            Target_Format = wd_XMLTemplateMacroEnabled
            Path_ToSave = Replace_Extension(Path_ToSave, ".dotm")

        Case "dotm"
            Target_Format = wd_Template97
            Path_ToSave = Replace_Extension(Path_ToSave, ".dot")

        Case Else: Exit Function

    End Select
     '`````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    Obj_Document.SaveAs2 FileName:=Path_ToSave, FileFormat:=Target_Format

    If Not CBool(Err.Number) Then Convert_WordExtension = Path_ToSave
    On Error GoTo 0
    '````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------'
End Function
'========================================================================'


'=========================================================================================='
Private Function Convert_PowerPointExtension( _
                ByRef Obj_Presentation As Object, _
                ByVal Path_ToSave As String _
        ) As String
'------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim File_Extension As String, Target_Format As Long
    '``````````````````````````````````````````````````'
    Const pp_Presentation      As Long = 1&
    Const pp_PresentationMacro As Long = 24&
    Const pp_Template          As Long = 4&
    Const pp_TemplateMacro     As Long = 29&
    Const pp_Show              As Long = 3&
    Const pp_ShowMacro         As Long = 25&
    Const pp_AddIn             As Long = 5&
    Const pp_AddInMacro        As Long = 28&
    '``````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    Convert_PowerPointExtension = vbNullString
    '`````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    File_Extension = LCase$(Right$(Obj_Presentation.Name, 4))

    Select Case File_Extension

        Case ".ppt"
            Target_Format = pp_PresentationMacro
            Path_ToSave = Replace_Extension(Path_ToSave, ".pptm")

        Case "pptm"
            Target_Format = pp_Presentation
            Path_ToSave = Replace_Extension(Path_ToSave, ".ppt")

        Case ".pot"
            Target_Format = pp_TemplateMacro
            Path_ToSave = Replace_Extension(Path_ToSave, ".potm")

        Case "potm"
            Target_Format = pp_Template
            Path_ToSave = Replace_Extension(Path_ToSave, ".pot")

        Case ".pps"
            Target_Format = pp_ShowMacro
            Path_ToSave = Replace_Extension(Path_ToSave, ".ppsm")

        Case "ppsm"
            Target_Format = pp_Show
            Path_ToSave = Replace_Extension(Path_ToSave, ".pps")

        Case ".ppa"
            Target_Format = pp_AddInMacro
            Path_ToSave = Replace_Extension(Path_ToSave, ".ppam")

        Case "ppam"
            Target_Format = pp_AddIn
            Path_ToSave = Replace_Extension(Path_ToSave, ".ppa")

        Case Else: Exit Function

    End Select
    '`````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    Obj_Presentation.SaveAs FileName:=Path_ToSave, FileFormat:=Target_Format
    '```````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    If Not CBool(Err.Number) Then Convert_PowerPointExtension = Path_ToSave
    On Error GoTo 0
    '```````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------'
End Function
'=========================================================================================='


'=========================================================================================='
Private Function Replace_Extension( _
                 ByRef File_Path As String, _
                 ByRef New_Extension As String _
        ) As String
'------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````````````````````'
    If InStrRev(File_Path, ".") > 0 Then
        Replace_Extension = Left$(File_Path, InStrRev(File_Path, ".") - 1) & New_Extension
    Else
        Replace_Extension = File_Path & New_Extension
    End If
    '``````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------'
End Function
'=========================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'=============================================================================================================='
Private Function VBProject_Hooking_DialogBoxParam() As Boolean
'--------------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````'
    Dim OS_Inx As Byte, Temp_Bytes(0 To 1) As Byte, Ptr_Address As LongPtr
    Dim Size_Region     As LongPtr, Ptr_OldProtect  As LongPtr
    Dim Trampoline_Size As LongPtr, Relative_Offset As Long
    Dim Error_Message   As String
    '`````````````````````````````````````````````````````````````````````'
    Const PAGE_EXECUTE_READ = &H20, PAGE_EXECUTE_READWRITE = &H40
    Const MEM_COMMIT = &H1000, MEM_RESERVE = &H2000
    '`````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    VBProject_Hooking_DialogBoxParam = False
    If Load_TrampolineAddress <> 0 Then Exit Function
    '````````````````````````````````````````````````'

    '```````````````````````````````'
    #If Win64 Then
        OS_Inx = 1: Size_Region = 12
    #Else
        OS_Inx = 0: Size_Region = 12
    #End If
    '```````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    Glb_DefaultAddress = Hack_WinAPI_Intercept("user32.dll", "DialogBoxParamA")
    If Glb_DefaultAddress = 0 Then
        Error_Message = "Не найдена функция API для проведения ин-лайн хука!"
        Show_ErrorMessage_Immediate Error_Message, "Не найден адрес функции!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If VirtualProtect(ByVal Glb_DefaultAddress, Size_Region, PAGE_EXECUTE_READWRITE, Ptr_OldProtect) = 0 Then
        Error_Message = "Не удается модифицировать права доступа к участку виртуальной памяти! " & _
                        "С большей долей вероятностью память уже имеет полные права доступа!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка модификации прав доступа!"
        Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    MoveMemory ByVal VarPtr(Temp_Bytes(0)), ByVal Glb_DefaultAddress, 1 + OS_Inx

    #If Win64 Then
        If Temp_Bytes(0) = &H48 And Temp_Bytes(1) = &HB8 Then
            Exit Function
        End If
    #Else
       If Temp_Bytes(OS_Inx) = &HB8 Then Exit Function
    #End If
    '```````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    MoveMemory ByVal VarPtr(Glb_Default_Bytes(0)), ByVal Glb_DefaultAddress, Size_Region

    ' // ToDo: Требуется оставить для теста
    ' //       Trampoline_Size = Size_Region + IIf(OS_Inx, 12, 5)

    Trampoline_Size = Size_Region + IIf(OS_Inx, 12, 8)

    Glb_TrampolineAddress = VirtualAlloc(0, Trampoline_Size, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If Glb_TrampolineAddress = 0 Then Exit Function

    MoveMemory ByVal Glb_TrampolineAddress, ByVal VarPtr(Glb_Default_Bytes(0)), Size_Region


    If Glb_TrampolineAddress <> 0 Then Save_TrampolineState


    #If Win64 Then
        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region), &H48, 1
        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 1), &HB8, 1
        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 2), ByVal VarPtr(Glb_DefaultAddress + 12), 8

        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 10), &HFF, 1
        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 11), &HE0, 1
    #Else

        Relative_Offset = (Glb_DefaultAddress + 12) - (Glb_TrampolineAddress + Size_Region + 5)

    ' // ToDo: Требуется оставить для теста
    ' //       MoveMemory ByVal (Glb_TrampolineAddress + Size_Region), &HE9, 1
    ' //       MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 1), Relative_Offset, 4

        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region), &H90, 3
        MoveMemory ByVal (Glb_TrampolineAddress + Size_Region + 3), &HE9, 1
    #End If
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    Ptr_Address = Get_PtrFunction(AddressOf Hack_DialogBoxParam)

    If CBool(OS_Inx) Then Glb_Hooking_Bytes(0) = &H48
    Glb_Hooking_Bytes(OS_Inx) = &HB8: OS_Inx = OS_Inx + 1

    MoveMemory ByVal VarPtr(Glb_Hooking_Bytes(OS_Inx)), ByVal VarPtr(Ptr_Address), 4 * OS_Inx
    Glb_Hooking_Bytes(OS_Inx + 4 * OS_Inx) = &HFF: Glb_Hooking_Bytes(OS_Inx + 4 * OS_Inx + 1) = &HE0

    Call MoveMemory(ByVal Glb_DefaultAddress, ByVal VarPtr(Glb_Hooking_Bytes(0)), Size_Region)
    VBProject_Hooking_DialogBoxParam = True

    Call VirtualProtect(ByVal Glb_DefaultAddress, Size_Region, PAGE_EXECUTE_READ, Ptr_OldProtect)
    '```````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------'
End Function
'=============================================================================================================='


'========================================================================================='
Private Function Hack_WinAPI_Intercept( _
                 ByRef WinAPI_Dll As String, _
                 ByRef WinAPI_Function As String _
        ) As LongPtr
'-----------------------------------------------------------------------------------------'
    Hack_WinAPI_Intercept = GetProcAddress(GetModuleHandleA(WinAPI_Dll), WinAPI_Function)
'-----------------------------------------------------------------------------------------'
End Function
'========================================================================================='


'==============================================================================='
Private Function Hack_DialogBoxParam( _
                 ByVal hInstance As LongPtr, _
                 ByVal pTemplateName As LongPtr, _
                 ByVal hWndParent As LongPtr, _
                 ByVal lpDialogFunc As LongPtr, _
                 ByVal dwInitParam As LongPtr _
        ) As Integer
'-------------------------------------------------------------------------------'

    '````````````````````````````````'
    Dim Trampoline_Address As LongPtr
    '````````````````````````````````'

    '````````````````````````````````````````````````````````````````'
    If pTemplateName = &HFE6 Then
        Hack_DialogBoxParam = &H1
    Else
        Trampoline_Address = Glb_TrampolineAddress

        If Trampoline_Address = 0& Then
            Trampoline_Address = Load_TrampolineAddress()
            If Trampoline_Address <> 0& Then
                Glb_TrampolineAddress = Trampoline_Address
            Else
                Hack_DialogBoxParam = &H1
                Application.DisplayAlerts = False: Application.Quit
                Exit Function
            End If
        End If

        Hack_DialogBoxParam = CInt(CallDlgBxParam( _
                                       Glb_TrampolineAddress, _
                                       hInstance, _
                                       pTemplateName, _
                                       hWndParent, _
                                       lpDialogFunc, _
                                       dwInitParam _
                                  ))
    End If
    '````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


'============================================='
Private Function Get_PtrFunction( _
                 ByVal Ptr_Value As LongPtr _
        ) As LongPtr
'---------------------------------------------'
    Get_PtrFunction = Ptr_Value
'---------------------------------------------'
End Function
'============================================='


'================================================================='
Private Function Save_TrampolineState()
'-----------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````'
    With ThisWorkbook
        On Error Resume Next
            .CustomDocumentProperties(CDP_HOOK_ACTIVE).Delete
            .CustomDocumentProperties(CDP_TRAMPOLINE_ADDR).Delete
            .CustomDocumentProperties(CDP_DEFAULT_ADDR).Delete
            .CustomDocumentProperties(CDP_DEFAULT_BYTES).Delete
        On Error GoTo 0

        .CustomDocumentProperties.Add _
            Name:=CDP_TRAMPOLINE_ADDR, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeFloat, _
            Value:=CLngPtr(Glb_TrampolineAddress)

        .CustomDocumentProperties.Add _
            Name:=CDP_HOOK_ACTIVE, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, _
            Value:=True

        .CustomDocumentProperties.Add _
            Name:=CDP_DEFAULT_ADDR, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeFloat, _
            Value:=CLngPtr(Glb_DefaultAddress)

        .CustomDocumentProperties.Add _
            Name:=CDP_DEFAULT_BYTES, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=StrConv(Glb_Default_Bytes, vbUnicode)
    End With
    '`````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------'
End Function
'================================================================='


'=================================================================================='
Private Function Load_TrampolineAddress() As LongPtr
'----------------------------------------------------------------------------------'

    '``````````````````````'
    Dim CDP_Value As String
    '``````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    If Not CBool(ThisWorkbook.CustomDocumentProperties(CDP_HOOK_ACTIVE).Value) Then
        Load_TrampolineAddress = 0
        On Error GoTo 0: Exit Function
    End If
    On Error GoTo 0
    '``````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    CDP_Value = ThisWorkbook.CustomDocumentProperties(CDP_TRAMPOLINE_ADDR).Value
    #If x64_Soft Then
        Load_TrampolineAddress = CLngPtr(CDP_Value)
    #Else
        Load_TrampolineAddress = CLng(CDP_Value)
    #End If
    '```````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------'
End Function
'=================================================================================='


'=============================================================='
Private Function Clear_TrampolineState()
'--------------------------------------------------------------'

    '`````````````````````````````````````````````````````````'
    On Error Resume Next
    With ThisWorkbook
        .CustomDocumentProperties(CDP_HOOK_ACTIVE).Delete
        .CustomDocumentProperties(CDP_TRAMPOLINE_ADDR).Delete
        .CustomDocumentProperties(CDP_DEFAULT_ADDR).Delete
        .CustomDocumentProperties(CDP_DEFAULT_BYTES).Delete
    End With
    On Error GoTo 0
    '`````````````````````````````````````````````````````````'

'--------------------------------------------------------------'
End Function
'=============================================================='


'=========================================================================================================='
Private Function Release_Trampoline()
'----------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````'
    Dim Trampoline_Address As LongPtr
    Dim defaultAddr        As LongPtr
    Dim defaultBytes()     As Byte
    Dim Ptr_OldProtect     As LongPtr
    Dim bytesLength        As Long
    Dim dblAddr            As Double
    Dim defaultBytesStr    As String
    '`````````````````````````````````'
    Const PAGE_EXECUTE_READ = &H20
    Const PAGE_EXECUTE_READWRITE = &H40
    Const MEM_RELEASE = &H8000

    '`````````````````````````````````````````'
    Trampoline_Address = Glb_TrampolineAddress
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````'
    If Trampoline_Address = 0 Then
        Trampoline_Address = Load_TrampolineAddress()
        If Trampoline_Address = 0 Then Exit Function

        On Error Resume Next
        dblAddr = ThisWorkbook.CustomDocumentProperties(CDP_DEFAULT_ADDR).Value
        defaultBytesStr = ThisWorkbook.CustomDocumentProperties(CDP_DEFAULT_BYTES).Value
        On Error GoTo 0

        If dblAddr = 0 Or defaultBytesStr = "" Then Exit Function
        defaultAddr = CLngPtr(dblAddr)
        defaultBytes = StrConv(defaultBytesStr, vbFromUnicode)
    Else
        defaultAddr = Glb_DefaultAddress
        defaultBytes = Glb_Default_Bytes
    End If
    '````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````'
    If defaultAddr <> 0 And Not IsEmpty(defaultBytes) Then
        bytesLength = UBound(defaultBytes) - LBound(defaultBytes) + 1

        If VirtualProtect(ByVal defaultAddr, bytesLength, PAGE_EXECUTE_READWRITE, Ptr_OldProtect) <> 0 Then
            MoveMemory ByVal defaultAddr, ByVal VarPtr(defaultBytes(LBound(defaultBytes))), bytesLength
            VirtualProtect ByVal defaultAddr, bytesLength, PAGE_EXECUTE_READ, Ptr_OldProtect
        End If
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````'
    If Trampoline_Address <> 0 Then
        VirtualFree Trampoline_Address, 0, MEM_RELEASE
        Glb_TrampolineAddress = 0
    End If
    '``````````````````````````````````````````````````'

    '`````````````````````````'
    Call Clear_TrampolineState
    '`````````````````````````'

'----------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================='


'==========================================================================================='
Private Function VBProject_Restore_DialogBoxParam()
'-------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````'
    Dim Size_Region As LongPtr, Inx As Long, i As Long
    '`````````````````````````````````````````````````'

    '```````````````````'
    #If Win64 Then
        Size_Region = 12
    #Else
        Size_Region = 12
    #End If
    '```````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    For i = 0 To UBound(Glb_Default_Bytes, 1): Inx = Inx + Glb_Default_Bytes(i): Next i

    If CBool(Inx) Then
        Call MoveMemory(ByVal Glb_Ptr_Func, ByVal VarPtr(Glb_Default_Bytes(0)), Size_Region)
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------'
End Function
'==========================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'======================================================================================================='
Private Function VBProject_Hooking_DialogBoxParam_RAM() As Boolean
'-------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````'
    Dim OS_Inx As Byte, Temp_Bytes(0 To 1) As Byte
    Dim Size_Region As LongPtr, Ptr_OldProtect As LongPtr, Ptr_Address As LongPtr

    Const PAGE_EXECUTE_READWRITE = &H40
    '````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    #If Win64 Then
        OS_Inx = 1: Size_Region = 12
    #Else
        OS_Inx = 0: Size_Region = 12
    #End If
    '```````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    VBProject_Hooking_DialogBoxParam_RAM = False
    Glb_Ptr_Func = Hack_WinAPI_Intercept("user32.dll", "DialogBoxParamA")
    '```````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    If VirtualProtect(ByVal Glb_Ptr_Func, Size_Region, PAGE_EXECUTE_READWRITE, Ptr_OldProtect) <> 0 Then

        MoveMemory ByVal VarPtr(Temp_Bytes(0)), ByVal Glb_Ptr_Func, 1 + OS_Inx
        If Temp_Bytes(OS_Inx) = &HB8 Then Exit Function

        MoveMemory ByVal VarPtr(Glb_Default_Bytes(0)), ByVal Glb_Ptr_Func, Size_Region
        Ptr_Address = Get_PtrFunction(AddressOf Hack_DialogBoxParam_RAM)

        If CBool(OS_Inx) Then Glb_Hooking_Bytes(0) = &H48
        Glb_Hooking_Bytes(OS_Inx) = &HB8: OS_Inx = OS_Inx + 1

        MoveMemory ByVal VarPtr(Glb_Hooking_Bytes(OS_Inx)), ByVal VarPtr(Ptr_Address), 4 * OS_Inx
        Glb_Hooking_Bytes(OS_Inx + 4 * OS_Inx) = &HFF: Glb_Hooking_Bytes(OS_Inx + 4 * OS_Inx + 1) = &HE0

        Call MoveMemory(ByVal Glb_Ptr_Func, ByVal VarPtr(Glb_Hooking_Bytes(0)), Size_Region)
        VBProject_Hooking_DialogBoxParam_RAM = True

    End If
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


'====================================================================================='
Private Function Hack_DialogBoxParam_RAM( _
                 ByVal hInstance As LongPtr, _
                 ByVal pTemplateName As LongPtr, _
                 ByVal hWndParent As LongPtr, _
                 ByVal lpDialogFunc As LongPtr, _
                 ByVal dwInitParam As LongPtr _
        ) As Integer
'-------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    If pTemplateName = &HFE6 Then
        Hack_DialogBoxParam_RAM = &H1
    Else
        VBProject_Restore_DialogBoxParam
        Hack_DialogBoxParam_RAM = _
        DialogBoxParam(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam)
    End If
    '`````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------'
End Function
'====================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'===================================================================================='
Private Function Convert_URLDecode( _
                 ByVal Text As String _
        ) As String
'------------------------------------------------------------------------------------'

    '```````````````````````````````````````````'
    Dim Start_Pos  As Long, Current_Pos As Long
    Dim First_Hex  As String
    Dim Third_Hex  As String
    Dim Second_Hex As String
    Dim Char_Code  As Long, Hex_Value   As Long
    Dim result   As String
    '```````````````````````````````````````````'
    Const Search_Char     As String = "%"
    Const Search_CharLength As Long = 1&
    '```````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Current_Pos = 1: Start_Pos = InStr(Current_Pos, Text, Search_Char)
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    Do While Start_Pos > 0

        If Current_Pos < Start_Pos Then
            result = result & Mid$(Text, Current_Pos, Start_Pos - Current_Pos)
        End If

        Select Case UCase$(Mid$(Text, Start_Pos + Search_CharLength, 1))
            Case "U"
                First_Hex = Mid$(Text, Start_Pos + Search_CharLength + 1, 4)
                Char_Code = CLng("&H" & First_Hex)
                result = result & ChrW$(Char_Code)
                Current_Pos = Start_Pos + 6

            Case "E"
                First_Hex = Mid$(Text, Start_Pos + Search_CharLength, 2)
                Hex_Value = CLng("&H" & First_Hex)

                If Hex_Value < &H80 Then
                    result = result & Chr$(Hex_Value)
                    Current_Pos = Start_Pos + 3
                Else
                    Second_Hex = Mid$(Text, Start_Pos + 3 + Search_CharLength, 2)
                    Third_Hex = Mid$(Text, Start_Pos + 6 + Search_CharLength, 2)

                    Char_Code = (CLng("&H" & First_Hex) And &HF) * 2 ^ 12 _
                             Or (CLng("&H" & Second_Hex) And &H3F) * 2 ^ 6 _
                             Or (CLng("&H" & Third_Hex) And &H3F)

                    If Char_Code < 0 Then Char_Code = Char_Code + 65536
                    result = result & ChrW$(Char_Code)
                    Current_Pos = Start_Pos + 9
                End If

            Case Else
                First_Hex = Mid$(Text, Start_Pos + Search_CharLength, 2)
                Hex_Value = CLng("&H" & First_Hex)

                If Hex_Value < &HC0 Then
                    result = result & Chr$(Hex_Value)
                    Current_Pos = Start_Pos + 3
                Else
                    Second_Hex = Mid$(Text, Start_Pos + 3 + Search_CharLength, 2)
                    Char_Code = (Hex_Value - &HC0) * &H40 + CLng("&H" & Second_Hex)
                    result = result & ChrW$(Char_Code)
                    Current_Pos = Start_Pos + 6
                End If
        End Select

        Start_Pos = InStr(Current_Pos, Text, Search_Char)
    Loop
    '````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If Current_Pos <= Len(Text) Then result = result & Mid$(Text, Current_Pos)
    '`````````````````````````````````````````````````````````````````````````'
    Convert_URLDecode = result
    '`````````````````````````'

'------------------------------------------------------------------------------------'
End Function
'===================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'==========================================================================================================='
Private Function ZIP_Archive_UnPack( _
                 ByRef zip_FilePath As String, _
                 ByRef tmp_FullName As String _
        ) As Boolean
'-----------------------------------------------------------------------------------------------------------'

    '``````````````'
    Dim T As Double
    '``````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````'
    With CreateObject("Shell.Application")
        .Namespace(CStr(zip_FilePath)).CopyHere .Namespace(CStr(tmp_FullName)).Items

        T = Timer

        On Error Resume Next
            Do Until .Namespace(CStr(tmp_FullName)).Items.Count = .Namespace(CStr(zip_FilePath)).Items.Count
                DoEvents: While Timer < T + 0.001: Wend
            Loop
        On Error GoTo 0
    End With
    '```````````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------'
End Function
'==========================================================================================================='


'============================================================================'
Private Function ZIP_Archive_Create( _
                 ByRef Path_FullName As String _
        ) As Boolean
'----------------------------------------------------------------------------'

    '``````````````````````````'
    Dim ZIP_Signature As String
    '``````````````````````````'

    '```````````````````````````````````````````````````````````````'
    ZIP_Archive_Create = False
    If Not Dir(Path_FullName) = vbNullString Then Kill Path_FullName
    '```````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    ZIP_Signature = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Open Path_FullName For Output As #1: Print #1, ZIP_Signature: Close #1
    '`````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````'
    If Not Dir(Path_FullName) = vbNullString Then ZIP_Archive_Create = True
    '``````````````````````````````````````````````````````````````````````'

    '``````````````````'
    Dir Environ$("TEMP")
    '``````````````````'

'----------------------------------------------------------------------------'
End Function
'============================================================================'


'==============================================================================='
Private Sub Remove_ItemXML( _
            ByRef Text_XML As String, _
            ByRef Item_XML As String _
        )
'-------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````'
    Dim Existence_Item As Long
    Dim Inx_Ending     As Long
    Dim Result_String  As String, Found_Item As String
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    If Not Left$(Item_XML, 1) = "<" Then Item_XML = "<" & Item_XML
    '`````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    Existence_Item = InStr(Text_XML, Item_XML)

    If CBool(Existence_Item) Then
        Inx_Ending = InStr(Existence_Item, Text_XML, "/>") + 2
        Found_Item = Mid$(Text_XML, Existence_Item, Inx_Ending - Existence_Item)
        Text_XML = Replace(Text_XML, Found_Item, vbNullString)
    End If
    '```````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Sub
'==============================================================================='


'================================================================================================='
Private Sub FileSystem_Select_FilesInExplorer( _
            ByRef Source_Directory As String, _
            ByRef Files_Name As Variant _
        )
'-------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````````````````'
    Dim Coll_Windows As Object, Item_Coll As Variant, SelectedItem, File_Name
    Dim Item_Window As Object, Item_Folder As Object, Item_File As Object, T As Double
    '`````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(Source_Directory) Then Exit Sub
    '`````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````'
    With CreateObject("Shell.Application")

        .Open CVar(Source_Directory): GoSub Time_Wait: Set Coll_Windows = .Windows

        If Right$(Source_Directory, 1) = "\" Then
            Source_Directory = Left$(Source_Directory, Len(Source_Directory) - 1)
        End If

        For Each Item_Window In Coll_Windows

            Set Item_Folder = Item_Window.Document

            If StrComp(Item_Folder.folder.Self.Path, Source_Directory, vbTextCompare) = 0 Then
                For Each SelectedItem In Item_Folder.SelectedItems
                    Item_Folder.SelectItem SelectedItem, &H0
                Next SelectedItem

                On Error Resume Next

                With CreateObject("Scripting.FileSystemObject")
                    For Each Item_Coll In Files_Name
                        File_Name = .GetFileName(Item_Coll)
                        Set Item_File = Item_Folder.folder.Items.Item(File_Name)
                        Item_Folder.SelectItem Item_File, &H19
                    Next Item_Coll
                End With

                On Error GoTo 0: Exit For
            End If

        Next Item_Window

    End With
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    Set Item_File = Nothing:   Set Item_Folder = Nothing
    Set Item_Window = Nothing: Set Coll_Windows = Nothing

    Exit Sub
    '````````````````````````````````````````````````````'

'```````````````````````````````````````````````````'
Time_Wait:

    T = Timer: DoEvents: While Timer < T + 0.5: Wend

    Return
'```````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Sub
'================================================================================================='


'==============================================================================='
Private Function Get_AllMetadata_File( _
                 ByRef File_FullName As String _
        ) As FileSystem_MetaData_File
'-------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````'
    Dim File_Name     As String, File_TempName  As String
    Dim File_BaseName As String, File_TypeName  As String
    Dim tmp_FullName  As String, tmp_FolderName As String
    '````````````````````````````````````````````````````'
    Dim tmp_Full_FolderName As String
    '````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    With CreateObject("Scripting.FileSystemObject")
        File_Name = .GetFileName(File_FullName)
        File_TempName = Left$(.GetTempName(), 8)
        File_BaseName = .GetBaseName(File_FullName)
        File_TypeName = .GetExtensionName(File_FullName)
    End With
    '``````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    tmp_FullName = Environ$("Temp") & "\" & File_BaseName & "_" & _
                                            File_TempName & "_." & File_TypeName

    tmp_FolderName = File_TempName & "_" & Format(Now, "DD.MM.YYYY") & " (" & _
                                           Format(Now, "HH-MM-SS") & ")"

    tmp_Full_FolderName = Environ$("TEMP") & "\" & tmp_FolderName
    '```````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Get_AllMetadata_File.MD_File_Name = File_Name
    Get_AllMetadata_File.MD_File_BaseName = File_BaseName
    Get_AllMetadata_File.MD_File_TempName = tmp_FullName
    Get_AllMetadata_File.MD_File_TypeName = File_TypeName
    Get_AllMetadata_File.MD_Folder_TempName = tmp_FolderName
    Get_AllMetadata_File.MD_Folder_TempDirectory = tmp_Full_FolderName
    '`````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


'============================================================================================================'
Private Function Get_SemanticDataType( _
                 ByRef Data As Variant _
        ) As MSOffice_Type_ContentFormat_Sec
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

    '``````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_VarType
        Case vbBoolean:              Get_SemanticDataType = MSDI_Bool_Sec:        Exit Function
        Case vbEmpty:                Get_SemanticDataType = MSDI_Empty_Sec:       Exit Function
        Case vbNull:                 Get_SemanticDataType = MSDI_Null_Sec:        Exit Function
        Case vbError, vbObjectError: Get_SemanticDataType = MSDI_Undefined_Sec:   Exit Function
        Case vbUserDefinedType:      Get_SemanticDataType = MSDI_UserDefType_Sec: Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If (Data_VarType And vbArray) And (Not Data_TypeName = "Range") Then
        Get_SemanticDataType = MSDI_Array_Sec: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_TypeName
        Case "String"
            If Len(Data) = 0 Then Get_SemanticDataType = MSDI_NullString_Sec: Exit Function
            If IsNumeric(Data) Then Get_SemanticDataType = MSDI_String_Sec:   Exit Function
            Data_Type = Get_Directory_Type(CStr(Data))
            
            Select Case Data_Type
                Case DirType_File:     Get_SemanticDataType = MSDI_File_Sec:                  Exit Function
                Case DirType_Folder:   Get_SemanticDataType = MSDI_Folder_Sec:                Exit Function
                Case DirType_NotFound: Get_SemanticDataType = MSDI_NonExistent_Directory_Sec: Exit Function
                Case DirType_Invalid:
                    On Error Resume Next
                    VBProject_Name = Application.VBE.ActiveVBProject.Name

                    If Not Err.Number = 0 Then
                        Error_Message = "Невозможно проверить тип данных (Компонент это или VB-Проект)! " & _
                                        "Тип данных будет определен как строка (String)!"
                        Call Show_ErrorMessage_Immediate(Error_Message, "Проблема идентификации типов")
                        Get_SemanticDataType = MSDI_String_Sec: On Error GoTo 0: Exit Function
                    End If

                    If VBProject_Name = Data Then
                        Get_SemanticDataType = MSDI_VBProject_Sec
                        On Error GoTo 0: Exit Function
                    End If

                    Err.Clear
                    VBComponent_Name = Application.VBE.ActiveVBProject.VBComponents.Item(Data).Name
                    
                    If Len(VBComponent_Name) > 0 And Err.Number = 0 Then
                        Get_SemanticDataType = MSDI_VBComponent_Sec
                        On Error GoTo 0: Exit Function
                    End If

                    For Each Obj_VBProjects In Application.VBE.VBProjects
                        If Obj_VBProjects.Name = Data Then
                            Get_SemanticDataType = MSDI_VBProject_Sec:   On Error GoTo 0: Exit Function
                        End If

                        Err.Clear: Set Obj_VBComponent = Obj_VBProjects.VBComponents.Item(Data)
                        If (Err.Number = 0) And (Not Obj_VBComponent Is Nothing) Then
                            If Not Len(Obj_VBComponent.Name) = 0 Then
                                Set Obj_VBComponent = Nothing
                                Get_SemanticDataType = MSDI_VBComponent_Sec: On Error GoTo 0: Exit Function
                            End If
                        End If
                        Set Obj_VBComponent = Nothing
                    Next Obj_VBProjects

                    Get_SemanticDataType = MSDI_String_Sec: On Error GoTo 0: Exit Function
            End Select

        Case "Byte":       Get_SemanticDataType = MSDI_Int8_Sec:             Exit Function
        Case "Integer":    Get_SemanticDataType = MSDI_Int16_Sec:            Exit Function
        Case "Long":       Get_SemanticDataType = MSDI_Int32_Sec:            Exit Function
        Case "LongLong":   Get_SemanticDataType = MSDI_Int64_Sec:            Exit Function
        Case "Single":     Get_SemanticDataType = MSDI_FloatPoint32_Sec:     Exit Function
        Case "Currency":   Get_SemanticDataType = MSDI_FloatPoint64Cur_Sec:  Exit Function
        Case "Double":     Get_SemanticDataType = MSDI_FloatPoint64Dbl_Sec:  Exit Function
        Case "Decimal":    Get_SemanticDataType = MSDI_FloatPoint112_Sec:    Exit Function
        Case "Date":       Get_SemanticDataType = MSDI_Date_Sec:             Exit Function
        Case "Range":      Get_SemanticDataType = MSDI_Range_Sec:            Exit Function
        Case "Dictionary": Get_SemanticDataType = MSDI_Dictionary_Sec:       Exit Function
        Case "Collection": Get_SemanticDataType = MSDI_Collection_Sec:       Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    If IsObject(Data) Then
        If Not Data Is Nothing Then
            Get_SemanticDataType = MSDI_Object_Sec:  Exit Function
        Else
            Get_SemanticDataType = MSDI_Nothing_Sec: Exit Function
        End If
    Else
        Get_SemanticDataType = MSDI_Undefined_Sec:   Exit Function
    End If
    '`````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'=============================================================================================================='
Private Sub Normalize_DataByType( _
            ByRef FileSystem_FilePath As Variant, _
            ByVal File_Type As MSOffice_Type_ContentFormat_Sec, _
            ByRef Count_Items As Long _
        )
'--------------------------------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Source_Array, i  As Long
    Dim tmp_Val, Inx_UB1 As Long
    Dim Type_ArrayDimensions
    Dim Error_Message As String
    '```````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case File_Type
        Case MSDI_Range_Sec, MSDI_Array_Sec
            Source_Array = FileSystem_FilePath

            If TypeName(Source_Array) = "String" Then
                tmp_Val = FileSystem_FilePath:    ReDim FileSystem_FilePath(1 To 1)
                FileSystem_FilePath(1) = tmp_Val: Count_Items = 1: Exit Sub
            End If

            Type_ArrayDimensions = IsArray_Dimensional(Source_Array)

            Select Case Type_ArrayDimensions

                Case Array_OneDimensional, Array_TwoDimensional
                    Inx_UB1 = UBound(Source_Array, 1)
                    If LBound(Source_Array, 1) = 0 Then Inx_UB1 = Inx_UB1 + 1

                    On Error Resume Next: ReDim FileSystem_FilePath(1 To Inx_UB1)

                    If Err.Number = 10 Then
                        Error_Message = "Работа со статическим массивом не поддерживается!"
                        Show_ErrorMessage_Immediate Error_Message, "Невозможно обработать массив", , True
                        On Error GoTo 0: Exit Sub
                    End If

                    On Error GoTo 0

                    If Type_ArrayDimensions = Array_OneDimensional Then
                        For i = LBound(Source_Array, 1) To UBound(Source_Array, 1)
                            If Not IsObject(Source_Array(i)) Then
                                Count_Items = Count_Items + 1
                                FileSystem_FilePath(Count_Items) = Source_Array(i)
                            End If
                        Next i
                    Else
                        For i = LBound(Source_Array, 1) To UBound(Source_Array, 1)
                            If Not IsObject(Source_Array(i, 1)) Then
                                Count_Items = Count_Items + 1
                                FileSystem_FilePath(Count_Items) = Source_Array(i, 1)
                            End If
                        Next i
                    End If

                Case Array_MultiDimensional
                    Error_Message = "Размерность массива выше 2-х мерной не поддерживается!"
                    Show_ErrorMessage_Immediate Error_Message, "Превышение допустимой размерности массива"
                    Exit Sub

                Case Array_NotInitialized
                    Error_Message = "Переданный массив на обработку не проинициализирован (является пустым)!"
                    Show_ErrorMessage_Immediate Error_Message, "Массив не проинициализирован"
                    Exit Sub

                Case Else
                    Error_Message = "Невозможно обработать переданный тип массива!"
                    Show_ErrorMessage_Immediate Error_Message, "Код типа данных - " & File_Type
                    Exit Sub

            End Select

        Case MSDI_File_Sec, MSDI_Folder_Sec
            tmp_Val = FileSystem_FilePath:     Count_Items = 1
            ReDim FileSystem_FilePath(1 To 1): FileSystem_FilePath(1) = tmp_Val

        Case MSDI_NonExistent_Directory_Sec
            Error_Message = "Директория или файл не найдены на текущем компьютере!"
            Show_ErrorMessage_Immediate Error_Message, FileSystem_FilePath
            Exit Sub

        Case Else
            Error_Message = "Невозможно обработать переданный тип данных!"
            Show_ErrorMessage_Immediate Error_Message, "Код типа данных - " & File_Type
            Exit Sub

    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------'
End Sub
'=============================================================================================================='


'==============================================================================='
Private Function IsArray_Dimensional( _
                 ByRef Assumed_Array As Variant _
        ) As MSOffice_Type_ArrayDimensions
'-------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````'
    Dim Dimensional_Count As Long, Array_Bound As Long
    '```````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    If Not IsArray(Assumed_Array) Then
        IsArray_Dimensional = Array_Error: Exit Function
    End If
    '```````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    On Error Resume Next: Array_Bound = LBound(Assumed_Array, 1)

    If Err.Number = 0 Then
        Array_Bound = UBound(Assumed_Array, 1)

        If Array_Bound = -1 Then
            IsArray_Dimensional = Array_NotInitialized: Exit Function
        End If

        Dimensional_Count = Dimensional_Count + 1
        Array_Bound = UBound(Assumed_Array, 2)

        If Err.Number = 0 Then
            Dimensional_Count = Dimensional_Count + 1
            Array_Bound = UBound(Assumed_Array, 3)
            If Err.Number = 0 Then Dimensional_Count = Dimensional_Count + 1
        End If
    End If

    On Error GoTo 0
    '````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Select Case Dimensional_Count
        Case 0:    IsArray_Dimensional = Array_NotInitialized
        Case 1:    IsArray_Dimensional = Array_OneDimensional
        Case 2:    IsArray_Dimensional = Array_TwoDimensional
        Case Else: IsArray_Dimensional = Array_MultiDimensional
    End Select
    '``````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


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


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================'
Private Function Create_WorkSheet_InExcel( _
                 ByRef Matrix_Data As Variant, _
                 ByRef Table_Headers As Variant, _
                 ByRef Table_Name As String, _
                 ByRef Table_ColumnWidth As Variant, _
                 ByRef WS_Name As String, _
                 ByRef WS_Code_Name As String _
        ) As Boolean
'----------------------------------------------------------------------------'

    '```````````````````'
    Dim Obj_WS As Object
    '```````````````````'

    '````````````````````````````````````````````````````````````````````````'
    Call Delete_Sheet_Name(WS_Name): Call Delete_Sheet_CodeName(WS_Code_Name)
    '````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    Set Obj_WS = Build_TableFromData(Matrix_Data, Table_Headers, _
                                     Table_Name, Table_ColumnWidth, _
                                     Application)
    '````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    If Obj_WS Is Nothing Then Create_WorkSheet_InExcel = False: Exit Function
    '````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    With Obj_WS
        .Name = WS_Name
        Application.VBE.ActiveVBProject. _
                    VBComponents(.codeName) _
                   .Name = WS_Code_Name
    End With

    Set Obj_WS = Nothing: Create_WorkSheet_InExcel = True
    '````````````````````````````````````````````````````'

'----------------------------------------------------------------------------'
End Function
'============================================================================'


'====================================================================='
Private Function Create_Process_ExcelApplication( _
                 ByRef Matrix_Data As Variant, _
                 ByRef Table_Headers As Variant, _
                 ByRef Table_Name As String, _
                 ByRef Table_ColumnWidth As Variant, _
                 ByRef WS_Name As String _
        ) As Boolean
'---------------------------------------------------------------------'

    '````````````````````````````````````````````````````````'
    Dim Obj_SDI As Object, Obj_WB As Object, Obj_WS As Object
    '````````````````````````````````````````````````````````'
    Const xl_Maximized As Long = -4137
    '`````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Set Obj_SDI = CreateObject("Excel.Application")

    With Obj_SDI

        Set Obj_WB = .Workbooks.Add: Set Obj_WS = Obj_WB.Worksheets(1)
        Call Build_TableFromData(Matrix_Data, Table_Headers, _
                                 Table_Name, Table_ColumnWidth, _
                                 Obj_SDI, Obj_WS)
        Obj_WS.Name = WS_Name
        .Visible = True: .WindowState = xl_Maximized

    End With
    '`````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If Obj_WS Is Nothing Then
        Create_Process_ExcelApplication = False: Exit Function
    Else
        Create_Process_ExcelApplication = True
    End If
    '`````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Set Obj_WS = Nothing: Set Obj_WB = Nothing: Set Obj_SDI = Nothing
    '`````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------'
End Function
'====================================================================='


'============================================================================='
Private Function Build_TableFromData( _
                 ByRef Matrix_Data As Variant, _
                 ByRef Table_Headers As Variant, _
                 Optional ByRef Table_Name As String, _
                 Optional ByRef Table_ColumnWidth As Variant = vbNullString, _
                 Optional ByRef Obj_SDI As Object = Nothing, _
                 Optional ByRef Obj_WS As Object = Nothing _
        ) As Object
'-----------------------------------------------------------------------------'

    '`````````````````````````````````````````````````'
    Dim Inx_UB1 As Long, Inx_UB2 As Long, N, i As Long
    '`````````````````````````````````````````````````'
    '````````````````````````````````````````````````````````````````````'
    Const xl_ThemeColorLight1 As Long = 2, xl_ThemeColorDark1 As Long = 1
    Const xl_EdgeBottom As Long = 9, xl_Center As Long = -4108
    Const xl_Continuous As Long = 1, xl_Medium As Long = -4138
    Const xl_R1C1       As Long = -4150
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    If Obj_SDI.Name = "Microsoft Excel" Then
        Obj_SDI.ScreenUpdating = False
        If Obj_WS Is Nothing Then
            On Error Resume Next
                Set Obj_WS = Obj_SDI.ThisWorkbook.Worksheets.Add
                If CBool(Err.Number) Then
                    Set Build_TableFromData = Obj_WS
                    Exit Function
                End If
            On Error GoTo 0
        End If
    End If

    On Error Resume Next
    Inx_UB1 = UBound(Matrix_Data, 1)
    If CBool(Err.Number) Then Inx_UB1 = 1: Err.Clear

    Inx_UB2 = UBound(Matrix_Data, 2)
    If CBool(Err.Number) Then Inx_UB2 = 1: Err.Clear
    On Error GoTo 0
    '``````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    With Obj_WS

        .Cells(1, 1) = "DATA"
        .Cells(1, 2) = Table_Name
        .Cells(2, 1).Resize(1, 4).FormulaR1C1 = Table_Headers

        With .Range("A1:E1")
            With .Font
                .Name = "Algerian"
                .Size = 20
                .Bold = True
                .ThemeColor = xl_ThemeColorLight1
            End With

            With .Interior
                .ThemeColor = xl_ThemeColorLight1
                .TintAndShade = 0.15
            End With
        End With

        With .Range("A1")
            .HorizontalAlignment = xl_Center
            .VerticalAlignment = xl_Center

            With .Font
                .ThemeColor = xl_ThemeColorDark1
            End With
        End With

        .Range("B1").Font.Color = RGB(255, 0, 0)

        With .Range("A2:E2")

            .HorizontalAlignment = xl_Center
            .VerticalAlignment = xl_Center

            With .Font
                .Bold = True
                .Name = "Futura Medium"
                .Size = 8
                .ThemeColor = xl_ThemeColorDark1
                .TintAndShade = -0.15
            End With

            With .Interior
                .ThemeColor = xl_ThemeColorLight1
                .TintAndShade = 0.35
            End With

            With .Borders(xl_EdgeBottom)
                .LineStyle = xl_Continuous
                .Color = RGB(255, 0, 0)
                .Weight = xl_Medium
            End With

            .AutoFilter

        End With

        .Rows(1).RowHeight = 28
        .Rows(2).RowHeight = 35

        .Range("A3").Resize(Inx_UB1, Inx_UB2).Value = Matrix_Data

        N = Split(.Cells(1, Inx_UB2).Address, "$")(1)

        If Not IsArray(Table_ColumnWidth) Then
            .Columns("A:" & N).EntireColumn.AutoFit
        Else
            For i = 1 To Inx_UB2
                .Columns(i).ColumnWidth = Table_ColumnWidth(i - 1)
            Next i
        End If

        With .Columns(Inx_UB2 + 1)
            .HorizontalAlignment = xl_Center
            .VerticalAlignment = xl_Center
            .ColumnWidth = 10
        End With

        .Cells(2, Inx_UB2 + 1).Resize(Inx_UB1 + 1, 1).Value = "-"

    End With
    '``````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    With Obj_SDI
        .GoTo Obj_WS.Cells(1, 1).Address(1, 1, xl_R1C1)
        .GoTo Obj_WS.Cells(3, 1).Address(1, 1, xl_R1C1)
        .ScreenUpdating = True
    End With
    '``````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Set Build_TableFromData = Obj_WS
    '```````````````````````````````'

'--------------------------------------------------------------------------'
End Function
'=========================================================================='


'======================================================'
Private Sub Delete_Sheet_Name( _
            ByRef Sheet_Name As String, _
            Optional Obj_WB As Object = Nothing _
        )
'------------------------------------------------------'

    '``````````````````````````````````````````````````'
    If Obj_WB Is Nothing Then Set Obj_WB = ThisWorkbook
    '``````````````````````````````````````````````````'

    '````````````````````````````````````````'
    On Error Resume Next
        Application.DisplayAlerts = False
        Obj_WB.Worksheets(Sheet_Name).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    '````````````````````````````````````````'

'------------------------------------------------------'
End Sub
'======================================================'


'======================================================================='
Private Sub Delete_Sheet_CodeName( _
            ByRef Sheet_CodeName As String, _
            Optional Obj_WB As Object = Nothing _
        )
'-----------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_WS As Object, tmp_WS
    '```````````````````````````'

    '``````````````````````````````````````````````````'
    If Obj_WB Is Nothing Then Set Obj_WB = ThisWorkbook
    '``````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    On Error Resume Next
    Set Obj_WS = Obj_WB.VBProject.VBComponents(Sheet_CodeName). _
                                                     CodeModule.Parent
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    If TypeName(Obj_WS) = "VBComponent" Then
        For Each tmp_WS In Obj_WB.Worksheets
            If tmp_WS.codeName = Sheet_CodeName Then
                Application.DisplayAlerts = False
                tmp_WS.Delete
                Application.DisplayAlerts = True
                Exit For
            End If
        Next tmp_WS
    End If

    On Error GoTo 0
    '````````````````````````````````````````````````'

'-----------------------------------------------------------------------'
End Sub
'======================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'================================================================================================================='
Private Function Create_UI_tmpForm() As Object
'-----------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Dim UI_UserForm As Object, UI_Control As Object, UI_Frame As Object, Inx_Replace As Long
    Dim UserForm_Property As UI_UserForm_Property, Control_Property As UI_Control_Property
    '```````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    UserForm_Property.UI_Caption = "UI_Progress_Bar"
    UserForm_Property.UI_Width = 618.75
    UserForm_Property.UI_Height = 85.5
    UserForm_Property.UI_ShowModal = False
    UserForm_Property.UI_ShowBorders = True
    UserForm_Property.UI_BackColor = vbBlack
    UserForm_Property.UI_BorderColor = vbBlack

    Set UI_UserForm = UI_Add_NewForm(UserForm_Property)
    '`````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    Control_Property.UI_Control_Type = Control_Label
    Control_Property.UI_Control_Name = "L_ProgressBar"
    Control_Property.UI_Control_Caption = "Обработано входных данных: [1] из [" & Glb_UI_Inx_Max & _
                                          "] _ Промежуточная обработка не требуется ||"
    Control_Property.UI_Control_FontSize = 9
    Control_Property.UI_Control_Left = 6
    Control_Property.UI_Control_Top = 6
    Control_Property.UI_Control_Width = 546
    Control_Property.UI_Control_Height = 24
    Control_Property.UI_Control_BackColor = vbBlack
    Control_Property.UI_Control_ForeColor = vbWhite
    Control_Property.UI_Control_FontBold = True

    Gbl_ProgressBar_Text = Control_Property.UI_Control_Caption

    Call UI_Add_NewControl(UI_UserForm, Control_Property): Control_Property = Clear_Control_Property()
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Control_Property.UI_Control_Type = Control_Frame
    Control_Property.UI_Control_Name = "Frm_1"
    Control_Property.UI_Control_FontSize = 9
    Control_Property.UI_Control_Left = 6
    Control_Property.UI_Control_Top = 24
    Control_Property.UI_Control_Width = 595
    Control_Property.UI_Control_Height = 23
    Control_Property.UI_Control_BackColor = vbBlack
    Control_Property.UI_Control_ForeColor = vbBlack

    Set UI_Frame = UI_Add_NewControl(UI_UserForm, Control_Property): Control_Property = Clear_Control_Property()
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````'
    Control_Property.UI_Control_Type = Control_CommandButton
    Control_Property.UI_Control_Name = "Butt_ProgressBar"
    Control_Property.UI_Control_Left = 1
    Control_Property.UI_Control_Top = 1
    Control_Property.UI_Control_Width = 590
    Control_Property.UI_Control_Height = 18
    Control_Property.UI_Control_BackColor = &HFFFF00
    Control_Property.UI_Control_ForeColor = &HFFFFFF

    Gbl_ProgressBar_Width = Control_Property.UI_Control_Width
    Control_Property.UI_Control_Width = 0

    Call UI_Add_NewControlToFrame(UI_Frame, Control_Property): Control_Property = Clear_Control_Property()
    '````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    UserForm_Property = Clear_UserForm_Property(): UserForms.Add(UI_UserForm.Name).Show
    Set Create_UI_tmpForm = UI_UserForm: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================='


'==========================================================='
Private Function Remove_UI_tmpForm( _
                 ByRef UI_UserForm As Variant _
        ) As Boolean
'-----------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim Obj_tmpForm As Object, Source_UserForm
    '``````````````````````````````````````````'

    '```````````````````````````````````````````````````````'
    For Each Source_UserForm In VBA.UserForms
        If Source_UserForm.Name = UI_UserForm.Name Then
            Source_UserForm.Hide
            Set Source_UserForm = Nothing: GoTo Remove_Form
        End If
    Next Source_UserForm
    '```````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````'
Remove_Form:

    With Application.VBE.ActiveVBProject
        While Obj_tmpForm Is Nothing
            Set Obj_tmpForm = .VBComponents(UI_UserForm.Name)
            .VBComponents.Remove Obj_tmpForm: DoEvents
        Wend
    End With

    Set Obj_tmpForm = Nothing: Set UI_UserForm = Nothing
'```````````````````````````````````````````````````````````'

'-----------------------------------------------------------'
End Function
'==========================================================='


'==============================================================================='
Private Function Clear_Control_Property() As UI_Control_Property:   End Function
Private Function Clear_UserForm_Property() As UI_UserForm_Property: End Function
'==============================================================================='


'============================================================================================================'
Private Function UI_Add_NewForm( _
                 ByRef UserForm_Property As UI_UserForm_Property _
        ) As Object
'------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim Obj_UI_Form As Object
    Set Obj_UI_Form = Application.VBE.ActiveVBProject.VBComponents.Add(3)
    '````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Len(UserForm_Property.UI_Name) = 0 Then
        UserForm_Property.UI_Name = Replace(CreateObject("Scripting.FileSystemObject").GetTempName, ".", "_")
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    With Obj_UI_Form

        .Properties("Name") = "vbFrm_" & UserForm_Property.UI_Name
        .Properties("Caption") = UserForm_Property.UI_Caption
        .Properties("Width") = UserForm_Property.UI_Width
        .Properties("Height") = UserForm_Property.UI_Height
        .Properties("ShowModal") = UserForm_Property.UI_ShowModal
        .Properties("BackColor") = UserForm_Property.UI_BackColor
        .Properties("BorderColor") = UserForm_Property.UI_BorderColor

        If Not Len(UserForm_Property.UI_CodeModule) = 0 Then
            With .CodeModule
                .InsertLines .CountOfLines + 1, UserForm_Property.UI_CodeModule
            End With
        End If

    End With
    '``````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````'
    Set UI_Add_NewForm = Obj_UI_Form
    '```````````````````````````````'

'------------------------------------------------------------------------------------------------------------'
End Function
'============================================================================================================'


'=================================================================================================================='
Private Function UI_Add_NewControl( _
                 ByRef UI_UserForm As Object, _
                 ByRef Control_Property As UI_Control_Property _
        ) As Object
'------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_UI_Control As Object
    '```````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Control_Property.UI_Control_Type
        Case Control_Label:         Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.Label.1")
        Case Control_TextBox:       Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.TextBox.1")
        Case Control_ComboBox:      Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.ComboBox.1")
        Case Control_ListBox:       Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.ListBox.1")
        Case Control_CheckBox:      Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.CheckBox.1")
        Case Control_OptionButton:  Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.OptionButton.1")
        Case Control_ToggleButton:  Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.ToggleButton.1")
        Case Control_Frame:         Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.Frame.1")
        Case Control_CommandButton: Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.CommandButton.1")
        Case Control_TabStrip:      Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.TabStrip.1")
        Case Control_MultiPage:     Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.MultiPage.1")
        Case Control_ScrollBar:     Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.ScrollBar.1")
        Case Control_SpinButton:    Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.SpinButton.1")
        Case Control_Image:         Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.Image.1")
        Case Control_RefEdit:       Set Obj_UI_Control = UI_UserForm.Designer.Controls.Add("Forms.RefEdit.1")
        Case Else:                  Set UI_Add_NewControl = Nothing: Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````'
    On Error Resume Next

    With Obj_UI_Control
        .Name = Control_Property.UI_Control_Name
        .Caption = Control_Property.UI_Control_Caption
        .Font.Size = Control_Property.UI_Control_FontSize
        .Font.Bold = Control_Property.UI_Control_FontBold
        .Left = Control_Property.UI_Control_Left
        .Top = Control_Property.UI_Control_Top
        .Width = Control_Property.UI_Control_Width
        .Height = Control_Property.UI_Control_Height
        .ForeColor = Control_Property.UI_Control_ForeColor
        .BackColor = Control_Property.UI_Control_BackColor
    End With
    '`````````````````````````````````````````````````````'

    '`````````````````````````````````````'
    Set UI_Add_NewControl = Obj_UI_Control
    '`````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================================='


'========================================================================================================================'
Private Function UI_Add_NewControlToFrame( _
                 ByRef UI_Frame As Object, _
                 ByRef Control_Property As UI_Control_Property _
        ) As Object
'------------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_UI_Control As Object
    '```````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Control_Property.UI_Control_Type
        Case Control_Label:         Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.Label.1", "btnInsideFrame")
        Case Control_TextBox:       Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.TextBox.1", "btnInsideFrame")
        Case Control_ComboBox:      Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.ComboBox.1", "btnInsideFrame")
        Case Control_ListBox:       Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.ListBox.1", "btnInsideFrame")
        Case Control_CheckBox:      Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.CheckBox.1", "btnInsideFrame")
        Case Control_OptionButton:  Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.OptionButton.1", "btnInsideFrame")
        Case Control_ToggleButton:  Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.ToggleButton.1", "btnInsideFrame")
        Case Control_CommandButton: Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.CommandButton.1", "btnInsideFrame")
        Case Control_TabStrip:      Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.TabStrip.1", "btnInsideFrame")
        Case Control_MultiPage:     Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.MultiPage.1", "btnInsideFrame")
        Case Control_ScrollBar:     Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.ScrollBar.1", "btnInsideFrame")
        Case Control_SpinButton:    Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.SpinButton.1", "btnInsideFrame")
        Case Control_Image:         Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.Image.1", "btnInsideFrame")
        Case Control_RefEdit:       Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.RefEdit.1", "btnInsideFrame")
        Case Else:                  Set UI_Add_NewControlToFrame = Nothing: Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````'
    On Error Resume Next

    With Obj_UI_Control
        .Name = Control_Property.UI_Control_Name
        .Caption = Control_Property.UI_Control_Caption
        .Font.Size = Control_Property.UI_Control_FontSize
        .Font.Bold = Control_Property.UI_Control_FontBold
        .Left = Control_Property.UI_Control_Left
        .Top = Control_Property.UI_Control_Top
        .Width = Control_Property.UI_Control_Width
        .Height = Control_Property.UI_Control_Height
        .ForeColor = Control_Property.UI_Control_ForeColor
        .BackColor = Control_Property.UI_Control_BackColor
    End With
    '`````````````````````````````````````````````````````'

    '````````````````````````````````````````````'
    Set UI_Add_NewControlToFrame = Obj_UI_Control
    '````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================================'


'========================================================================================================================'
Private Function Update_UI_tmpForm_ProgressBar_Upper( _
                 ByVal Current_Inx As Long, _
                 ByRef UI_tmpForm_Name As String _
        ) As Boolean
'------------------------------------------------------------------------------------------------------------------------'

    '````````````````````````'
    Dim UI_UserForm As Object
    '````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            If Not Current_Inx = Glb_UI_Inx_Max Then
                Gbl_ProgressBar_Text = Replace(Gbl_ProgressBar_Text, "[" & Current_Inx & "]", "[" & Current_Inx + 1 & "]")
                UI_UserForm.L_ProgressBar.Caption = Gbl_ProgressBar_Text
            End If
            UI_UserForm.Butt_ProgressBar.Width = Gbl_ProgressBar_Width / Glb_UI_Inx_Max * Current_Inx
            Update_UI_tmpForm_ProgressBar_Upper = True: Exit Function
        End If
    Next UI_UserForm

    Update_UI_tmpForm_ProgressBar_Upper = False
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================================'


'============================================================================================'
Private Function Update_UI_tmpForm_ProgressBar_Lower( _
                 ByVal Current_Inx As Long, _
                 ByVal Max_Inx As Long, _
                 ByRef UI_tmpForm_Name As String _
        ) As Boolean
'--------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````'
    Dim UI_UserForm As Object, tmp_Arr As Variant
    '````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    tmp_Arr = Split(Gbl_ProgressBar_Text, "_")
    tmp_Arr(1) = "  Промежуточная обработка - {" & Current_Inx & "} из {" & Max_Inx & "} ||"
    Gbl_ProgressBar_Text = Join(tmp_Arr, "_")
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            UI_UserForm.L_ProgressBar.Caption = Gbl_ProgressBar_Text
            If Current_Inx = Max_Inx Then
                tmp_Arr(1) = " Промежуточная обработка не требуется ||"
                Gbl_ProgressBar_Text = Join(tmp_Arr, "_")
                UI_UserForm.L_ProgressBar.Caption = Gbl_ProgressBar_Text
            End If
            Update_UI_tmpForm_ProgressBar_Lower = True: Exit Function
        End If
    Next UI_UserForm

    Update_UI_tmpForm_ProgressBar_Lower = False
    '```````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------'
End Function
'============================================================================================'


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
