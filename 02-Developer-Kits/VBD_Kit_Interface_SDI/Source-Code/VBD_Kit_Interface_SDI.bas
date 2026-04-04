Attribute VB_Name = "VBD_Kit_Interface_SDI"
' | ================================================================================================== | '
' | ________              ______ ________               ________             ________           _____  | '
' | ___  __ \_____ __________  /___  ___/___________    ___  __ \__________________(_)____________  /_ | '
' | __  / / /  __ `/_  ___/_  //_/____ \_  _ \  ___/    __  /_/ /_  ___/  __ \____  /_  _ \  ___/  __/ | '
' | _  /_/ // /_/ /_  /   _  ,<  ____/ //  __/ /__      _  ____/_  /   / /_/ /___  / /  __/ /__ / /_   | '
' | /_____/ \__,_/ /_/    /_/|_| /____/ \___/\___/______/_/     /_/    \____/___  /  \___/\___/ \__/   | '
' |                                              _/_____/                    /___/                     | '
' | ================================================================================================== | '

' +-[MODULE: VBD_Kit_Interface_SDI]----------------------------------------------------------+
' |                                                                                          |
' | [ENGINEER]: Zeus_0x01                                                                    |
' | [TELEGRAM]: @Zeus_0x01 (Public Name)                                                     |
' | [DESCRIPTION]: Реализация SDI_API для процессов Excel с "изолированным" исполнением кода |
' |                                                                                          |
' +------------------------------------------------------------------------------------------+

' // <copyright file="VBD_Kit_Interface_SDI.bas" division="DarkSec_Project">
' // (C) Copyright 2023 Zeus_0x01 "{AB9BC9C2-E7DC-4702-AAC1-6AED940B2947}"
' // </copyright>

'-------------------------------------------------------------------------'
' // Implemented Functionality (Реализованный функционал):
'    Forum -  https://script-coding.ru/threads/vbd_kit_interface_sdi.211/
'    GitHub - https://github.com/Cyber-Automation/XL_INTERNALS/
'-------------------------------------------------------------------------'
' // Release_Version (Версия компонента) - [01.01]
'-------------------------------------------------------------------------'

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

    ' - < Инкапсуляция компонента >
    #If BuildMode_PrivateModule Then
        Option Private Module
    #End If
' }
'============================================================================'

'---------------------------------------------------------------------------------'
Private Const GUID_VBComponent As String = "{AB9BC9C2-E7DC-4702-AAC1-6AED940B2947}"
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

'---------------------------------'
Public Enum Process_Priority
    Unknown_Priority = &HFFFFFFFF
    Low_Priority = &H40
    BelowNormal_Priority = &H4000
    Normal_Priority = &H20
    AboveNormal_Priority = &H8000
    High_Priority = &H80
    RealTime_Priority = &H100
End Enum
'---------------------------------'

'---------------------------------'
Public Enum Process_TerminateType
    Normal_ExlApp = &H0
    Emergency_WinAPI = &H1
End Enum
'---------------------------------'

'--------------------------------'
Public Enum Process_RestrictType
    Restrict_No = &H0
    Restrict_Excel = &H1
End Enum
'--------------------------------'

'-----------------------------'
Public Enum Process_Type
    Background_Process = &H0
    Application_Process = &H1
End Enum
'-----------------------------'

'--------------------------'
Public Enum Thread_State
    T_State_Resume = &H0
    T_State_Suspended = &H1
End Enum
'--------------------------'

'------------------------------------------'
Public Enum Thread_TypeCommand
    CommandThread_Initialize
    CommandThread_Terminate
    CommandThread_Run
    CommandThread_Halt
    CommandThread_Save
    CommandThread_Resume
    CommandThread_Suspended
    CommandThread_AbortExecution
    CommandThread_ChangePriority
    CommandThread_EfficiencyMode
    CommandThread_ProcessAffinity
    CommandThread_VBProject_AddComponent
    CommandThread_VBProject_RemoveComponent
    CommandThread_VBProject_UpdateComponent
End Enum
'------------------------------------------'

'------------------------------------------'
Private Enum Thread_StatusCommand
    StatusCommandThread_NeedAbort = &HFFEE
    StatusCommandThread_NotFound = &HFFFE
    StatusCommandThread_Error = &HFFFF
    StatusCommandThread_Initialize = &H0
    StatusCommandThread_Done = &H1
End Enum
'------------------------------------------'

'----------------------------------'
Private Enum Process_State
    P_State_ActiveExecution
    P_State_AbortExecution
    P_State_Idle
    P_State_Completed
    P_State_Suspended
    P_State_Debugging
    P_State_Terminate
    P_State_NonExistent
    P_State_ErrorReading
    P_State_Unresponsive
    P_State_ManuallyStopped
    P_State_UserTriggeredExecution
End Enum
'----------------------------------'

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

'--------------------------------------------'
Private Enum MSOffice_Type_ContentFormat_SDI
    MSDI_Undefined_SDI = &HFFFFFFFC
    MSDI_Null_SDI = &HFFFFFFFD
    MSDI_NullString_SDI = &HFFFFFFFE
    MSDI_Nothing_SDI = &HFFFFFFFF
    MSDI_Empty_SDI = &H0
    MSDI_Bool_SDI = &H1
    MSDI_Int8_SDI = &H2
    MSDI_Int16_SDI = &H4
    MSDI_Int32_SDI = &H8
    MSDI_Int64_SDI = &H10
    MSDI_FloatPoint32_SDI = &H20
    MSDI_FloatPoint64Cur_SDI = &H40
    MSDI_FloatPoint64Dbl_SDI = &H56
    MSDI_FloatPoint112_SDI = &H80
    MSDI_Date_SDI = &H100
    MSDI_String_SDI = &H200
    MSDI_Range_SDI = &H400
    MSDI_Array_SDI = &H800
    MSDI_Object_SDI = &H1000
    MSDI_Collection_SDI = &H2000
    MSDI_Dictionary_SDI = &H4000
    MSDI_File_SDI = &H8000
    MSDI_Folder_SDI = &H10000
    MSDI_VBComponent_SDI = &H20000
    MSDI_VBProject_SDI = &H40000
    MSDI_UserDefType_SDI = &H80000
    MSDI_NonExistent_Directory_SDI = &H100000
    
    ' {
        MSDI_Procedure_SDI = &H120000
        MSDI_AutoDetect_SDI = &HFFFFFFAD
    ' }
End Enum
'--------------------------------------------'

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

'------------------------------------------'
Private Enum WinAPI_ThreadAccessRights
    NT_THREAD_TERMINATE = &H1
    NT_THREAD_SUSPEND_RESUME = &H2
    NT_THREAD_GET_CONTEXT = &H8
    NT_THREAD_SET_CONTEXT = &H10
    NT_THREAD_QUERY_INFORMATION = &H40
    NT_THREAD_SET_INFORMATION = &H20
    NT_THREAD_SET_THREAD_TOKEN = &H80
    NT_THREAD_IMPERSONATE = &H100
    NT_THREAD_DIRECT_IMPERSONATION = &H200
    NT_THREAD_ALL_ACCESS = &H1F03FF
End Enum
'------------------------------------------'

'------------------------------------------------'
Private Enum WinAPI_ProcessAccessRights
    NT_PROCESS_TERMINATE = &H1
    NT_PROCESS_CREATE_THREAD = &H2
    NT_PROCESS_SET_SESSIONID = &H4
    NT_PROCESS_VM_OPERATION = &H8
    NT_PROCESS_VM_READ = &H10
    NT_PROCESS_VM_WRITE = &H20
    NT_PROCESS_DUP_HANDLE = &H40
    NT_PROCESS_CREATE_PROCESS = &H80
    NT_PROCESS_SET_QUOTA = &H100
    NT_PROCESS_SET_INFORMATION = &H200
    NT_PROCESS_QUERY_INFORMATION = &H400
    NT_PROCESS_SUSPEND_RESUME = &H800
    NT_PROCESS_QUERY_LIMITED_INFORMATION = &H1000
    NT_PROCESS_SET_LIMITED_INFORMATION = &H2000
    NT_PROCESS_ALL_ACCESS = &H1F0FFF
    STANDARD_RIGHTS_REQUIRED = &HF0000
    SYNCHRONIZE = &H100000
End Enum
'------------------------------------------------'

'--------------------------------------------------'
Private Enum WinAPI_CreateToolhelp32Snapshot_lFlags
    TH32CS_SNAPPROCESS = &H2
    TH32CS_SNAPTHREAD = &H4
    TH32CS_SNAPMODULE = &H8
    TH32CS_SNAPMODULE32 = &H10
    TH32CS_SNAPHEAPLIST = &H1
End Enum
'--------------------------------------------------'

'---------------------------------'
Private Type ProcessAffinity_Mask
    Affinity_Mask       As String
    Active_LogicalCores As Long
    Count_LogicalCores  As Long
End Type
'---------------------------------'

'----------------------------------------------------------------'
Private Type Process_Information
    Init_Process              As Boolean
    Handle_Process            As Long
    RAM_ID_Process            As String
    Process_IDentifier        As Long
    Executable_Component      As String
    Additional_Component()    As Variant
    Efficiency_Mode           As Boolean
    Priority_Level            As Process_Priority
    Process_Affinity          As ProcessAffinity_Mask
    Type_Executable_Component As MSOffice_Type_ContentFormat_SDI
    State_Process             As Process_State
End Type
'----------------------------------------------------------------'

'-----------------------------------------'
Public Type Process_Snapshot
    Init_Structure  As Boolean
    Count_Process   As Long
    List_PID()      As Long
    Processes_SDI() As Process_Information
End Type
'-----------------------------------------'

'-----------------------'
Private Type WinAPI_GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type
'-----------------------'

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

'------------------------------------------'
Private Type WinAPI_SystemInfo
    dwOemId                     As Long
    dwPageSize                  As Long
    lpMinimumApplicationAddress As LongPtr
    lpMaximumApplicationAddress As LongPtr
    dwActiveProcessorMask       As LongPtr
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    dwProcessorLevel            As Integer
    dwProcessorRevision         As Integer
End Type
'------------------------------------------'

'---------------------------------'
Private Type WinAPI_ThreadEntry32
    dwSize             As Long
    cntUsage           As Long
    th32ThreadID       As Long
    rh32OwnerProcessID As Long
    tpBasePri          As Long
    tpDeltaPri         As Long
    dwFlags            As Long
End Type
'---------------------------------'

'--------------------------------------'
Private Type WinAPI_ProcessEntry32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As LongPtr
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * 260
End Type
'--------------------------------------'

'------------------------------------------------'
Private Type WinAPI_Process_PowerThrottling_State
    Version     As Long
    ControlMask As Long
    StateMask   As Long
End Type
'------------------------------------------------'

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
#If Windows_NT And Has_PtrSafe Then  ' // Windows API (Kernel32.dll)

    Private Declare PtrSafe _
            Function CreateThread Lib "kernel32.dll" ( _
                     ByVal lpThreadAttributes As LongPtr, _
                     ByVal dwStackSize As Long, _
                     ByVal lpStartAddress As LongPtr, _
                     ByRef lpParameter As Any, _
                     ByVal dwCreationFlags As Long, _
                     ByRef lpThreadId As Long _
            ) As LongPtr

    Private Declare PtrSafe _
            Function OpenProcess Lib "kernel32.dll" ( _
                     ByVal dwDesiredAccess As Long, _
                     ByVal bInheritHandle As Long, _
                     ByVal dwProcessId As Long _
            ) As Long

    Private Declare PtrSafe _
            Function TerminateProcess Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByVal uExitCode As Long _
            ) As Long

    Private Declare PtrSafe _
            Function SetProcessInformation Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByVal ProcessInformationClass As Long, _
                     ByRef ProcessInformation As Any, _
                     ByVal ProcessInformationLength As Long _
            ) As Long

    Private Declare PtrSafe _
            Function SetPriorityClass Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByVal dwPriorityClass As Long _
            ) As Long

    Private Declare PtrSafe _
            Function GetPriorityClass Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Function SetProcessAffinityMask Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByVal dwProcessAffinityMask As LongPtr _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetProcessAffinityMask Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByRef lpProcessAffinityMask As LongPtr, _
                     ByRef lpSystemAffinityMask As LongPtr _
            ) As LongPtr

    Private Declare PtrSafe _
            Sub GetSystemInfo Lib "kernel32.dll" ( _
                ByRef lpSystemInfo As WinAPI_SystemInfo _
            )

    Private Declare PtrSafe _
            Function GetExitCodeProcess Lib "kernel32.dll" ( _
                     ByVal hProcess As LongPtr, _
                     ByRef lpExitCode As Long _
            ) As Long

    Private Declare PtrSafe _
            Function CloseHandle Lib "kernel32.dll" ( _
                     ByVal hObject As LongPtr _
            ) As Long

    Private Declare PtrSafe _
            Sub Sleep Lib "kernel32.dll" ( _
                ByVal dwMilliseconds As Long _
            )

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
            Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
                     ByVal lFlags As Long, _
                     ByVal lProcessId As Long _
            ) As Long

    Private Declare PtrSafe _
            Function Thread32First Lib "kernel32.dll" ( _
                     ByVal hSnapshot As Long, _
                     ByRef uProcess As WinAPI_ThreadEntry32 _
            ) As Long

    Private Declare PtrSafe _
            Function Thread32Next Lib "kernel32.dll" ( _
                     ByVal hSnapshot As Long, _
                     ByRef uProcess As WinAPI_ThreadEntry32 _
            ) As Long

    Private Declare PtrSafe _
            Function OpenThread Lib "kernel32.dll" ( _
                     ByVal dwDesiredAccess As Long, _
                     ByVal bInheritHandle As Boolean, _
                     ByVal dwThreadId As Long _
            ) As Long

    Private Declare PtrSafe _
            Function ResumeThread Lib "kernel32.dll" ( _
                     ByVal hThread As Long _
            ) As Integer

    Private Declare PtrSafe _
            Function SuspendThread Lib "kernel32.dll" ( _
                     ByVal hThread As Long _
            ) As Integer

    Private Declare PtrSafe _
            Function Process32First Lib "kernel32.dll" ( _
                     ByVal hSnapshot As LongPtr, _
                     ByRef lppe As WinAPI_ProcessEntry32 _
            ) As Long

    Private Declare PtrSafe _
            Function Process32Next Lib "kernel32.dll" ( _
                     ByVal hSnapshot As LongPtr, _
                     ByRef lppe As WinAPI_ProcessEntry32 _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'--------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then  ' // Windows API (User32.dll)

    Private Declare PtrSafe _
            Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
                     ByVal hWnd1 As LongPtr, _
                     ByVal hwnd2 As LongPtr, _
                     ByVal lpsz1 As String, _
                     ByVal lpsz2 As String _
            ) As Long

    Private Declare PtrSafe _
            Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" ( _
                     ByVal hWnd As Long, _
                     ByVal Msg As Long, _
                     ByVal wParam As LongPtr, _
                     ByVal lParam As LongPtr, _
                     ByVal fuFlags As Long, _
                     ByVal uTimeout As Long, _
                     ByRef lpdwResult As LongPtr _
            ) As LongPtr

    Private Declare PtrSafe _
            Function GetWindowThreadProcessId Lib "user32.dll" ( _
                     ByVal hWnd As LongPtr, _
                     ByRef lpdwProcessId As Long _
            ) As Long

    Private Declare PtrSafe _
            Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" ( _
                     ByVal hWnd As Long, _
                     ByVal lpString As String, _
                     ByVal cch As Long _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'---------------------------------------------------------------------------------------'

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

'----------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then  ' // Windows API (Psapi.dll)

    Private Declare PtrSafe _
            Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
                     ByVal hProcess As LongPtr, _
                     ByVal hModule As LongPtr, _
                     ByVal lpFileName As String, _
                     ByVal nSize As Long _
            ) As Long

#Else

    ' // Old MS Office or MacOS

#End If
'----------------------------------------------------------------------------------------'

'-----------------------------------------------'
Private Glb_Thread_HaltExecution      As Boolean
Private Glb_Thread_StateExecution     As Boolean
'------------------------------------------------------------------'
Private Glb_MSOffice_Type_Building    As MSOffice_Type_Building
Private Glb_MSOffice_Type_Application As MSOffice_Type_Application
'------------------------------------------------------------------'

'---------------------------------------------'
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const WM_NULL          As Long = &H0
Private Const STILL_ACTIVE     As Long = &H103
'---------------------------------------------'

'----------------------------------------------------------'
Private Const INVALID_HANDLE_VALUE As LongPtr = &HFFFFFFFF
'----------------------------------------------------------'

'-----------------------------------------------------------------------'
Private Const IDLE_PRIORITY_CLASS         As Long = Low_Priority
Private Const BELOW_NORMAL_PRIORITY_CLASS As Long = BelowNormal_Priority
Private Const NORMAL_PRIORITY_CLASS       As Long = Normal_Priority
Private Const ABOVE_NORMAL_PRIORITY_CLASS As Long = AboveNormal_Priority
Private Const HIGH_PRIORITY_CLASS         As Long = High_Priority
Private Const REALTIME_PRIORITY_CLASS     As Long = RealTime_Priority
'-----------------------------------------------------------------------'
Private Const PROCESS_POWER_THROTTLING_EXECUTION_SPEED  As Long = &H1
'-----------------------------------------------------------------------'

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

'--------------------------------------------'
Private Const TIMEOUT_PROCESS As Long = 4096&
'-------------------------------------------------------------------'
Private Const VBComponent_KernelName  As String = "modKernel_Thread"
Private Const VBComponent_RuntimeName As String = "RT_VB4I5C"
'-------------------------------------------------------------------'


'================================================================================================================================'
Public Function Process_Create( _
                ByRef SDIProcess_Snapshot As Process_Snapshot, _
                Optional ByVal Executable_Component As String = vbNullString, _
                Optional ByVal Additional_Component As String = vbNullString, _
                Optional ByVal Count_Process As Long = 1, _
                Optional ByVal Type_Process As Process_Type = Background_Process, _
                Optional ByVal Code_RunningWhenReady As Boolean = True, _
                Optional ByVal TerminateWhenExecuted As Boolean = True, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal dw_Reserved_2 As Long = 0&, _
                Optional ByVal dw_Reserved_3 As Long = 0& _
       ) As Boolean
'--------------------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_create-vbd_kit_interface_sdi-bas.277/
'````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - Persistence_Mode | Внедрить ли процесс в систему для автоматического запуска
' dw_Reserved_2 - Password_ESM     | Пароль для диспетчера пользовательских служб Excel (Только при "Persistence_Mode = True")
' dw_Reserved_3 - Password_SDI     | Пароль на процесс для диспетчера служб Excel (Только при "Persistence_Mode = True")
'--------------------------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````````'
    Dim SDI_NameComponent As String, SDI_Workbook As Object, SDI_VBComponent As Object
    Dim SDI_Process       As Object, Empty_SDIProcess_Snapshot     As Process_Snapshot
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Vector_ListComponent()       As String, Vector_stdModule_Code() As String
    Dim Vector_AdditionalComponent() As Variant, Vector_Components      As Variant
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Obj_VBComponent As Object, Obj_CheckComponent As Object, Obj_CodeModule As Object
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Dict_FullNameVBComponents As Object, Dict_NameVBComponents As Object
    Dim VBComponent_FullName      As String, VBComponent_ToCopy    As String
    Dim Type_ContentFormat        As MSOffice_Type_ContentFormat_SDI
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim VBComponent_Name As String, Clone_Component As Boolean, Type_VBComponent As Long
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Affinity_Mask  As String, Component_Code As String, Error_Message   As String
    Dim File_Extension As String, State_Process  As Process_State, Err_Number As Long
    '````````````````````````````````````````````````````````````````````````````````````'
    Dim Inx_Process As Long, PID As Long, Count_AdditionalComponent As Long
    Dim Inx_LB1 As Long, Inx_UB1 As Long, Inx As Long, i As Long, N As Long
    '````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````'
    Process_Create = False
    '`````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    SDI_NameComponent = Find_ModuleByGUID(GUID_VBComponent).Name
    '```````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Executable_Component = vbNullString Then

        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Выберите исполняемый компонент для нового процесса Excel (.bas)"
            .Filters.Clear: .Filters.Add "Executable_Component VBA", "*.bas"
            .AllowMultiSelect = False

            If .Show = -1 Then
                Executable_Component = .SelectedItems(1): Type_ContentFormat = MSDI_File_SDI
            Else
                Error_Message = "Процесс не может быть создан без исполняемого компонента"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                Exit Function
            End If

            If Executable_Component Like "*\" & VBComponent_KernelName & ".bas" Then
                Error_Message = "Процесс не может быть создан c именем исполняемого компонента " & _
                                "(modKernel_Thread)! Данное имя зарезервировано под компонент ядра процесса!"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                Exit Function
            ElseIf Executable_Component = "*\" & SDI_NameComponent & ".bas" Then
                Error_Message = "Процесс не может быть создан c именем исполняемого компонента " & _
                                "(VBD_Kit_Interface_SDI)! " & _
                                "Данное имя зарезервировано под компонент API для интерфейса SDI!"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                Exit Function
            End If
        End With

    Else

        If CBool(InStr(1, Executable_Component, "|")) Then
            Error_Message = "Процесс не может быть создан! Проверьте наименование компонента - " _
                            & Executable_Component
            Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
            Exit Function
        End If

        If Executable_Component = VBComponent_KernelName Then
            Error_Message = "Процесс не может быть создан c именем исполняемого компонента " & _
                            "(modKernel_Thread)! Данное имя зарезервировано под компонент ядра процесса!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
            Exit Function
        ElseIf Executable_Component = SDI_NameComponent Then
            Error_Message = "Процесс не может быть создан c именем исполняемого компонента " & _
                            "(VBD_Kit_Interface_SDI)! " & _
                            "Данное имя зарезервировано под компонент API для интерфейса SDI!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
            Exit Function
        End If

        Type_ContentFormat = Get_SemanticDataType(Executable_Component)

        If Not Type_ContentFormat = MSDI_VBComponent_SDI Then
            If Not Type_ContentFormat = MSDI_File_SDI Then
                On Error Resume Next
                VBComponent_Name = _
                Application.VBE.ActiveVBProject.VBComponents.Item(Executable_Component).Name
                If Len(VBComponent_Name) = 0 Then Type_ContentFormat = MSDI_Empty_SDI
                On Error GoTo 0
            End If
        End If

        Select Case Type_ContentFormat
            Case MSDI_VBComponent_SDI
                If Application.VBE.ActiveVBProject.VBComponents.Item(Executable_Component).Type <> 1& Then
                    Error_Message = _
                    "В качестве исполняемого файла допускаются только стандартные модули (StdModule) с разрешением "".bas"""
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                    Exit Function
                End If
            Case MSDI_File_SDI
                If Right$(Executable_Component, 4) <> ".bas" Then
                    Error_Message = _
                    "В качестве исполняемого файла допускаются только стандартные модули (StdModule) с разрешением "".bas"""
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                    Exit Function
                End If
            Case Else
                Error_Message = "Компонент """ & Executable_Component & _
                             """ не существует в текущем проекте VBA"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                Exit Function
        End Select

    End If
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Additional_Component = vbNullString Then
        Set Dict_NameVBComponents = CreateObject("Scripting.Dictionary")
        Dict_NameVBComponents(Executable_Component) = vbNullString

        Set Dict_FullNameVBComponents = CreateObject("Scripting.Dictionary")
        Dict_FullNameVBComponents(Executable_Component) = vbNullString

        Vector_Components = Application.Trim(Split(Additional_Component, "|"))

        For i = LBound(Vector_Components, 1) To UBound(Vector_Components, 1)
            VBComponent_FullName = Vector_Components(i)

            VBComponent_ToCopy = Mid$( _
                                     VBComponent_FullName, _
                                     InStrRev(VBComponent_FullName, "\") + 1 _
                                 )

            If VBComponent_ToCopy <> VBComponent_FullName Then
                VBComponent_ToCopy = Left$(VBComponent_ToCopy, _
                                     InStr(VBComponent_ToCopy, ".") - 1)
            End If

            If Not Dict_NameVBComponents.Exists(VBComponent_ToCopy) Then
                Dict_NameVBComponents(VBComponent_ToCopy) = vbNullString
                Dict_FullNameVBComponents(VBComponent_FullName) = vbNullString
            End If
        Next i

        Vector_Components = Dict_FullNameVBComponents.keys: ReDim Vector_ListComponent(1 To UBound(Vector_Components))

        For i = 1 To UBound(Vector_ListComponent)
            If Not Len(Vector_Components(i)) = 0 Then
                On Error Resume Next
                    Set Obj_CheckComponent = Application.VBE.ActiveVBProject.VBComponents(Vector_Components(i))
                    If Obj_CheckComponent Is Nothing Then
                        Select Case Get_SemanticDataType(Vector_Components(i))
                            Case MSDI_File_SDI
                            Case MSDI_NonExistent_Directory_SDI
                                On Error GoTo 0
                                Error_Message = "Дополнительный компонент не найден в файловой системе - " & Vector_Components(i)
                                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                                Exit Function
                            Case Else
                                On Error GoTo 0
                                Error_Message = "Дополнительный компонент не найден в текущем проекте - " & Vector_Components(i)
                                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                                Exit Function
                        End Select

                        File_Extension = Right$(Vector_Components(i), 4)
                        Select Case File_Extension
                            Case ".bas", ".cls", ".frm"
                            Case Else
                                Error_Message = _
                                "В качестве дополнительного компонента не допускаются файлы с разрешением " & _
                                File_Extension & ". Разрешённые форматы - "".bas"", "".cls"", "".frm"""
                                Show_ErrorMessage_Immediate Error_Message, "Ошибка создания процесса Excel"
                                Exit Function
                        End Select
                    End If
                    Set Obj_CheckComponent = Nothing
                On Error GoTo 0
                Vector_ListComponent(i) = Vector_Components(i)
            End If
        Next i

        Additional_Component = Join(Vector_ListComponent, "|")
    End If
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    With SDIProcess_Snapshot
        If .Init_Structure Then
            Inx_Process = .Count_Process + 1
            .Count_Process = .Count_Process + Count_Process
            ReDim Preserve .Processes_SDI(1 To .Count_Process)
            ReDim Preserve .List_PID(1 To .Count_Process)
        Else
            .Init_Structure = True: Inx_Process = 1
            .Count_Process = Count_Process
            ReDim .Processes_SDI(1 To .Count_Process)
            ReDim .List_PID(1 To .Count_Process)
        End If

        For i = Inx_Process To .Count_Process
            With .Processes_SDI(i)
                If Not Code_RunningWhenReady Then
                    .State_Process = P_State_Idle
                End If
                .Executable_Component = Executable_Component
                .Type_Executable_Component = Type_ContentFormat
                GoSub GS¦Initialize_Proccess: .Init_Process = True
                GoSub GS¦Copy_AdditionalComponent
                
                If Code_RunningWhenReady Then
                    SDI_Process.Run "'" & SDI_Workbook.FullName & "'!Thread_Initialize"
                    .State_Process = SDI_Workbook. _
                                     CustomDocumentProperties("State_Process")
                    If .State_Process <> P_State_ActiveExecution Then
                        Error_Message = "Не удается запустить процедуру ""Main""! Проверьте компонент - " & _
                                         Executable_Component
                        Show_ErrorMessage_Immediate Error_Message, "Ошибка запуска процедуры"
                        GoSub GS¦Terminate_Proccess: Exit Function
                    End If
                End If

                If Type_Process = Application_Process Then SDI_Process.Visible = True

                .Efficiency_Mode = False
                .Priority_Level = Threads_Change_Priority(.Process_IDentifier, Unknown_Priority)

                Affinity_Mask = Threads_Get_ProcessAffinity(.Process_IDentifier, Restrict_Excel)
                If Len(Affinity_Mask) > 0 Then
                    With .Process_Affinity
                        .Affinity_Mask = Split(Affinity_Mask, "_")(0)
                         Affinity_Mask = Split(Affinity_Mask, "_")(1)
                        .Active_LogicalCores = Split(Affinity_Mask, "|")(0)
                        .Count_LogicalCores = Split(Affinity_Mask, "|")(1)
                    End With
                End If

                On Error Resume Next
                    N = LBound(Vector_AdditionalComponent, 1)
                    If Err.Number = 0 Then
                        Sorting_VBComponents Vector_AdditionalComponent, _
                                             LBound(Vector_AdditionalComponent), _
                                             UBound(Vector_AdditionalComponent)
                    End If
                On Error GoTo 0

                .Additional_Component = Vector_AdditionalComponent
            End With
        Next i
    End With
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````'
    Process_Create = True: Exit Function
    '```````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````'
GS¦Initialize_Proccess:

    With SDIProcess_Snapshot.Processes_SDI(i)

        Set SDI_Process = CreateObject("Excel.Application")
        Set SDI_Workbook = SDI_Process.Workbooks.Add

        State_Process = .State_Process:
        Call GetWindowThreadProcessId(SDI_Process.hWnd&, PID)

        .Handle_Process = SDI_Process.hWnd
        .RAM_ID_Process = SDI_Workbook.FullName
        .Process_IDentifier = PID

        If UBound(SDIProcess_Snapshot.List_PID) < i Then
            ReDim Preserve SDIProcess_Snapshot.List_PID(1 To i)
        End If

        SDI_Process.DisplayAlerts = False
        SDIProcess_Snapshot.List_PID(i) = PID
        GoSub GS¦Build_modKernel_Thread: GoSub GS¦Copy_ExecutableComponent

        With SDI_Workbook.CustomDocumentProperties

            .Add "Hwnd_XlMain", _
                  False, msoPropertyTypeString, Application.hWnd
            .Add "Thread_RunningWhenReady", _
                  False, msoPropertyTypeBoolean, Code_RunningWhenReady
            .Add "Thread_WaitingAfterExecution", _
                  False, msoPropertyTypeBoolean, TerminateWhenExecuted
            .Add "Executable_Component", _
                  False, msoPropertyTypeString, Executable_Component
            .Add "State_Process", _
                  False, msoPropertyTypeNumber, State_Process
            .Add "PID", _
                  False, msoPropertyTypeNumber, PID
            .Add "ExecuteCommand_ErrorCode", _
                  False, msoPropertyTypeNumber, StatusCommandThread_Initialize

            If Code_RunningWhenReady Then
                .Add "ExecuteCommand_Status", _
                      False, msoPropertyTypeNumber, CommandThread_Initialize
            Else
                .Add "ExecuteCommand_Status", _
                      False, msoPropertyTypeNumber, CommandThread_Halt
            End If

        End With

    End With

    Return
'``````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````'
GS¦Terminate_Proccess:

    If SDIProcess_Snapshot.Count_Process = 1 Then
        SDIProcess_Snapshot = Empty_SDIProcess_Snapshot
    Else
        With SDIProcess_Snapshot
            .Count_Process = .Count_Process - 1
            ReDim Preserve .List_PID(1 To .Count_Process)
        End With

        With Empty_SDIProcess_Snapshot
            .Count_Process = SDIProcess_Snapshot.Count_Process
            .Init_Structure = True: .List_PID = SDIProcess_Snapshot.List_PID
            ReDim .Processes_SDI(1 To .Count_Process)

            For N = 1 To .Count_Process
                With .Processes_SDI(N)
                    .Additional_Component = _
                     SDIProcess_Snapshot.Processes_SDI(N).Additional_Component

                    .Efficiency_Mode = _
                     SDIProcess_Snapshot.Processes_SDI(N).Efficiency_Mode

                    .Executable_Component = _
                     SDIProcess_Snapshot.Processes_SDI(N).Executable_Component

                    .Handle_Process = _
                     SDIProcess_Snapshot.Processes_SDI(N).Handle_Process

                    .Init_Process = _
                     SDIProcess_Snapshot.Processes_SDI(N).Init_Process

                    .Priority_Level = _
                     SDIProcess_Snapshot.Processes_SDI(N).Priority_Level

                    .Process_Affinity = _
                     SDIProcess_Snapshot.Processes_SDI(N).Process_Affinity

                    .Process_IDentifier = _
                     SDIProcess_Snapshot.Processes_SDI(N).Process_IDentifier

                    .RAM_ID_Process = _
                     SDIProcess_Snapshot.Processes_SDI(N).RAM_ID_Process

                    .State_Process = _
                     SDIProcess_Snapshot.Processes_SDI(N).State_Process

                    .Type_Executable_Component = _
                     SDIProcess_Snapshot.Processes_SDI(N).Type_Executable_Component
                End With
            Next N
        End With

        SDIProcess_Snapshot = Empty_SDIProcess_Snapshot
    End If
    
    SDI_Workbook.Close False: Set SDI_Workbook = Nothing
    SDI_Process.Quit:         Set SDI_Process = Nothing
    
    Return
'```````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GS¦Copy_ExecutableComponent:

    VBComponent_ToCopy = Executable_Component: GoSub GS¦Build_VBComponent

    If Not Clone_Component Then
        GoSub GS¦Terminate_Proccess: On Error GoTo 0: Exit Function
    End If

    Return
'````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Copy_AdditionalComponent:

    If Additional_Component = vbNullString Then Return

    On Error Resume Next
        Count_AdditionalComponent = UBound(Vector_AdditionalComponent, 1)
    On Error GoTo 0

    If Count_AdditionalComponent = 0 Then
        If InStr(1, Additional_Component, "|") > 0 Then
            Vector_AdditionalComponent = Application.Trim(Split(Additional_Component, "|"))
        Else
            Vector_AdditionalComponent = Application.Trim(Array((Additional_Component)))
        End If

        Inx = 0
        Inx_LB1 = LBound(Vector_AdditionalComponent, 1)
        Inx_UB1 = UBound(Vector_AdditionalComponent, 1)

        For N = Inx_LB1 To Inx_UB1
            VBComponent_ToCopy = Vector_AdditionalComponent(N)
            Type_ContentFormat = Get_SemanticDataType(VBComponent_ToCopy)
            If Type_ContentFormat = MSDI_File_SDI Or Type_ContentFormat = MSDI_VBComponent_SDI Then
                Inx = Inx + 1: Vector_AdditionalComponent(Inx) = VBComponent_ToCopy
            End If
        Next N

        If Inx_UB1 > Inx Then
            ReDim Preserve Vector_AdditionalComponent(1 To Inx)
        End If
    End If

    If Inx > 0 Then
        For N = 1 To Inx
            VBComponent_ToCopy = Vector_AdditionalComponent(N)
            Type_ContentFormat = Get_SemanticDataType(VBComponent_ToCopy)
            GoSub GS¦Build_VBComponent
            If Not Clone_Component Then
                GoSub GS¦Terminate_Proccess: On Error GoTo 0: Exit Function
            End If
        Next N
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Build_VBComponent:

    Clone_Component = False

    If Type_ContentFormat = MSDI_File_SDI Then

        SDI_Workbook.VBProject.VBComponents.Import VBComponent_ToCopy
        VBComponent_FullName = VBComponent_ToCopy

        VBComponent_ToCopy = Mid$( _
                                  VBComponent_ToCopy, _
                                  InStrRev(VBComponent_ToCopy, "\") + 1 _
                             )

        VBComponent_ToCopy = Left$(VBComponent_ToCopy, _
                             InStr(VBComponent_ToCopy, ".") - 1)

        On Error Resume Next: Type_VBComponent = -1
        Type_VBComponent = SDI_Workbook.VBProject.VBComponents(VBComponent_ToCopy).Type
        On Error GoTo 0

        If Type_VBComponent = -1 Then
            Error_Message = "Структура компонента не определена! Проверьте компонент - " & _
                             VBComponent_ToCopy
            Show_ErrorMessage_Immediate Error_Message, "Ошибка добавления компонента"
            GoSub GS¦Terminate_Proccess: Exit Function
        End If

        If Type_VBComponent = 100 Then Type_VBComponent = 2
        SDI_Workbook.VBProject.VBComponents(VBComponent_ToCopy).Name = VBComponent_RuntimeName

        Set Obj_VBComponent = SDI_Workbook.VBProject.VBComponents(VBComponent_RuntimeName)
        Component_Code = Obj_VBComponent.CodeModule.Lines(1, Obj_VBComponent.CodeModule.CountOfLines)

        With SDI_Workbook.VBProject
            .VBComponents.Remove Obj_VBComponent
        End With

        GoSub GS¦Copy_VBComponent

    Else

        Component_Code = Empty
        Set Obj_VBComponent = ThisWorkbook.VBProject.VBComponents(VBComponent_ToCopy)
        Type_VBComponent = Obj_VBComponent.Type: If Type_VBComponent = 100 Then Type_VBComponent = 2

        If Not Obj_VBComponent.CodeModule.CountOfLines = 0 Then
            Component_Code = Obj_VBComponent.CodeModule.Lines(1, Obj_VBComponent.CodeModule.CountOfLines)
        End If

        On Error Resume Next
            Set Obj_CheckComponent = SDI_Workbook.VBProject.VBComponents(VBComponent_ToCopy)
            Err_Number = Err.Number
        On Error GoTo 0

        If Obj_CheckComponent Is Nothing Or Err_Number <> 0 Then
            GoSub GS¦Copy_VBComponent
        End If

    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Copy_VBComponent:

    If VBComponent_ToCopy = Executable_Component Or Executable_Component = VBComponent_FullName Then
        Component_Code = Component_Code & String$(3, vbNewLine) & Replace( _
                                   Get_Sub_InModule(SDI_VBComponent, Obj_CodeModule, _
                                                  "Update_KernelThread_Code"), _
                                                  "Private", "Public" _
                                   )
    End If

    With SDI_Workbook.VBProject
        .VBComponents.Add(Type_VBComponent).Name = VBComponent_ToCopy: Clone_Component = True
        .VBComponents(VBComponent_ToCopy).CodeModule.AddFromString Component_Code

        With .VBComponents(VBComponent_ToCopy).CodeModule
            If .CountOfLines > 0 Then
                Component_Code = .Lines(1, .CountOfLines)
                If Right$(Component_Code, 2) = "()" Then .DeleteLines (.CountOfLines)
                If InStr(1, Component_Code, "Option Explicit") = 1 Then .DeleteLines 1
                If .Lines(1, 1) = vbNullString Then .DeleteLines 1
            End If
        End With

    End With

    VBComponent_FullName = vbNullString

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````'
GS¦Injection_modKernel_Thread:

    With SDI_Workbook.VBProject
        .VBComponents.Add(1).Name = VBComponent_KernelName
        With .VBComponents(VBComponent_KernelName).CodeModule
            .AddFromString Component_Code
            Component_Code = .Lines(1, .CountOfLines)
            If Right$(Component_Code, 2) = "()" Then .DeleteLines (.CountOfLines)
            If InStr(1, Component_Code, "Option Explicit") = 1 Then .DeleteLines 1
            If .Lines(1, 1) = vbNullString Then .DeleteLines 1
        End With
    End With

    Component_Code = vbNullString

    Return
'``````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Build_modKernel_Thread:

    Set SDI_VBComponent = Find_ModuleByGUID(GUID_VBComponent): Set Obj_CodeModule = SDI_VBComponent.CodeModule
    ReDim Vector_stdModule_Code(1 To 19) As String

    Vector_stdModule_Code(1) = Get_Enums_InModule(Obj_CodeModule, "Process_State") & vbCrLf
    Vector_stdModule_Code(2) = Get_Enums_InModule(Obj_CodeModule, "MSOffice_Type_ContentFormat") & vbCrLf
    Vector_stdModule_Code(3) = Get_Enums_InModule(Obj_CodeModule, "Thread_StatusCommand") & vbCrLf
    Vector_stdModule_Code(4) = Replace( _
                               Get_Enums_InModule(Obj_CodeModule, "Thread_TypeCommand") & vbCrLf, _
                               "Public", "Private")

    Vector_stdModule_Code(5) = Get_Enums_InModule(Obj_CodeModule, "FileSystem_Directory_Type") & vbCrLf

    Vector_stdModule_Code(6) = "'-------------------------------------------------------------------'"
    Vector_stdModule_Code(7) = "Private Const VBComponent_KernelName  As String = """ & VBComponent_KernelName & """"
    Vector_stdModule_Code(8) = "Private Const VBComponent_RuntimeName As String = """ & VBComponent_RuntimeName & """"
    Vector_stdModule_Code(9) = "'-------------------------------------------------------------------'" & vbCrLf

    Vector_stdModule_Code(10) = Replace( _
                               Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, _
                               "Controller_Thread") & vbCrLf, "Private", "Public")

    Vector_stdModule_Code(11) = Get_Sub_InModule(SDI_VBComponent, Obj_CodeModule, "Thread_Initialize") & vbCrLf
    Vector_stdModule_Code(12) = Get_Sub_InModule(SDI_VBComponent, Obj_CodeModule, "Thread_Completed") & vbCrLf
    Vector_stdModule_Code(13) = Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, "Get_SemanticDataType") & vbCrLf
    Vector_stdModule_Code(14) = Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, "Get_Directory_Type") & vbCrLf
    Vector_stdModule_Code(15) = Get_Sub_InModule(SDI_VBComponent, Obj_CodeModule, "Deferred_ProcedureCall") & vbCrLf
    Vector_stdModule_Code(16) = Get_Sub_InModule(SDI_VBComponent, Obj_CodeModule, "Create_ModuleWithCode") & vbCrLf
    Vector_stdModule_Code(17) = Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, "IsModule_Exists") & vbCrLf
    Vector_stdModule_Code(18) = Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, "IsProcedure_Exists") & vbCrLf
    Vector_stdModule_Code(19) = Get_Function_InModule(SDI_VBComponent, Obj_CodeModule, "Show_ErrorMessage_Immediate") _
                                                                                                          & vbCrLf
    Component_Code = Join(Vector_stdModule_Code, vbCrLf): GoSub GS¦Injection_modKernel_Thread

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================================'


'=================================================================================================='
Public Function Process_Terminate( _
                ByRef SDIProcess_Snapshot As Process_Snapshot, _
                Optional ByVal Terminate_Type As Process_TerminateType = Normal_ExlApp, _
                Optional ByVal Process_ID As Long = -1, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal dw_Reserved_2 As Long = 0& _
       ) As Boolean
'--------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_terminate-vbd_kit_interface_sdi-bas.278/
'``````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - Disable_Persistence | Отключить у процесса возможность автоматического запуска
' dw_Reserved_2 - Password_ESM        | Пароль для диспетчера служб Excel (Опционально)
'--------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````'
    Dim Flag_Result  As Boolean, Count_SDIProcess As Long
    Dim Error_Message As String, i As Long
    '````````````````````````````````````````````````````'

    '````````````````````````'
    Process_Terminate = False
    '````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````'
    If Process_ID = -1 Then
        With SDIProcess_Snapshot
            Count_SDIProcess = .Count_Process: Flag_Result = True
            If Count_SDIProcess > 0 Then
                For i = Count_SDIProcess To 1 Step -1
                    Process_ID = .Processes_SDI(i).Process_IDentifier
                    GoSub GS¦Terminate_Process
                    If Process_Terminate Then
                        GoSub GS¦Update_Process
                    Else
                        Flag_Result = False
                    End If
                Next i
            End If
        End With
        Process_Terminate = Flag_Result
    Else
        GoSub GS¦Terminate_Process: GoSub GS¦Update_Process
    End If

    Exit Function
    '````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````'
GS¦Update_Process:

    If Process_Terminate Then
        If i = 0 Then
            Count_SDIProcess = SDIProcess_Snapshot.Count_Process
            For i = Count_SDIProcess To 1 Step -1
                If Process_ID = SDIProcess_Snapshot.List_PID(i) Then
                    Exit For
                End If
            Next i
        End If

        With SDIProcess_Snapshot.Processes_SDI(i)
            .Handle_Process = -1
            .Init_Process = False
            .Process_IDentifier = -1
            .State_Process = P_State_Terminate
            .Priority_Level = Unknown_Priority
            .Efficiency_Mode = False

            With .Process_Affinity
                .Active_LogicalCores = -1
                .Affinity_Mask = vbNullString
                .Count_LogicalCores = -1
            End With
        End With
    End If

    Return
'````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GS¦Terminate_Process:

    If Terminate_Type = Emergency_WinAPI Then
        Process_Terminate = Process_Kill(Process_ID, Restrict_Excel)
    Else
        Process_Terminate = Process_ExecuteCommand( _
                                           SDIProcess_Snapshot, _
                                           CommandThread_Terminate, Process_ID, , False _
                                    )
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'============================================================================================='
Public Function Process_Kill( _
                ByVal Process_Metric As Variant, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As Boolean
'---------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_kill-vbd_kit_interface_sdi-bas.281/
'---------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````'
    Dim Processes_PID    As Variant, i As Long
    Dim Flag_ProcessKill As Boolean, Count_Process As Long
    '`````````````````````````````````````````````````````'

    '```````````````````'
    Process_Kill = False
    '```````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    If IsNumeric(Process_Metric) Then
        If Process_Restrict = Restrict_Excel Then
            If Not UCase$(Get_ProcessPath_PID(Process_Metric)) Like "*EXCEL.EXE" Then
                Process_Kill = False: Exit Function
            End If
        End If
        Process_Kill = KillProcess_PID(Process_Metric)
    Else
        Processes_PID = Get_RunningProcesses_ByName(Process_Metric)
        Count_Process = UBound(Processes_PID, 1): Flag_ProcessKill = True
        If Not Count_Process = -1 Then
            For i = 1 To Count_Process
                Process_Kill = KillProcess_PID(Processes_PID(i))
                If Not Process_Kill Then Flag_ProcessKill = False
            Next i
            Process_Kill = Flag_ProcessKill
        End If
    End If
    '````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------'
End Function
'============================================================================================='


'=================================================================================================='
Public Function Process_Save( _
                ByRef SDIProcess_Snapshot As Process_Snapshot, _
                ByVal Process_ID As Long, _
                ByRef Save_Directory As String _
       ) As Variant
'--------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_save-vbd_kit_interface_sdi-bas.282/
'--------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Error_Message As String, Flag_PID As Boolean, i As Long
    '``````````````````````````````````````````````````````````'

    '```````````````````'
    Process_Save = False
    '```````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    If Save_Directory = vbNullString Then
        Error_Message = "Требуется указать директорию (папку) для сохранения данных процесса"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка сохранения процесса Excel"
        Exit Function
    Else
        If Not Get_Directory_Type(Save_Directory) = DirType_Folder Then
            Error_Message = "Указанный путь не является директорией - " & Save_Directory
            Show_ErrorMessage_Immediate Error_Message, "Ошибка сохранения процесса Excel"
            Exit Function
        End If
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    If Not Right$(Save_Directory, 1) = "\" Then Save_Directory = Save_Directory & "\"
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    With SDIProcess_Snapshot
        If .Count_Process = 0 Then
            Error_Message = "В структуре процессов нет ни одного процесса!"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка сохранения процесса Excel"
            Exit Function
        End If

        Flag_PID = False

        For i = 1 To .Count_Process
            If .List_PID(i) = Process_ID Then Flag_PID = True: Exit For
        Next i

        If Not Flag_PID Then
            Error_Message = "В структуре процессов нет указанного процесса (" & Process_ID & ")"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка сохранения процесса Excel"
            Exit Function
        End If
    End With
    '```````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````'
    Process_Save = Process_ExecuteCommand( _
                           SDIProcess_Snapshot, _
                           CommandThread_Save, _
                           Process_ID, _
                           Save_Directory, False _
                   )
    '``````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    If Process_Save Then
        Process_Save = Save_Directory & SDIProcess_Snapshot. _
                                        Processes_SDI(i).RAM_ID_Process & ".xlsx"
    End If
    '````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'=========================================================================================================='
Public Function Process_Get_State( _
                ByRef SDIProcess_Snapshot As Process_Snapshot, _
                Optional ByVal Process_ID As Long = -1, _
                Optional ByVal dw_Reserved_1 As Long = 0& _
       ) As Boolean
'----------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_save-vbd_kit_interface_sdi-bas.282/
'``````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - ToDo_Parameter | Зарезервировано без явного описания функционала
'----------------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````````'
    Dim COM_Application  As Object, PID As Long, RAM_ID As String
    Dim Affinity_Mask As Variant, Hwnd_Process As Long, Err_Number As Long
    '`````````````````````````````````````````````````````````````````````'
    Dim Flag_NotResponsive As Boolean, Error_Message   As String
    Dim State_Process As Process_State, VBProject_Mode As Long
    '`````````````````````````````````````````````````````````````````````'
    Dim Count_Process As Long, Inx_Process  As Long, i As Long
    '`````````````````````````````````````````````````````````````````````'

    '````````````````````````'
    Process_Get_State = False
    '````````````````````````'

    '`````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '`````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Count_Process = SDIProcess_Snapshot.Count_Process: If Count_Process = 0 Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    With SDIProcess_Snapshot
        If Not Process_ID = -1 Then
            For i = 1 To Count_Process
                If Process_ID = .List_PID(i) Then
                    Inx_Process = i: Exit For
                End If
            Next i

            If Inx_Process = -1 Then Exit Function Else GoSub GS¦Read_Process
        Else
            For i = Count_Process To 1 Step -1
                Inx_Process = i: GoSub GS¦Read_Process
            Next i

            If Inx_Process = -1 Then
                Error_Message = "В структуре процессов нет указанного процесса (" & Process_ID & ")"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка чтения процесса Excel"
            End If
        End If
    End With

    Process_Get_State = True

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Read_Process:

    With SDIProcess_Snapshot

        PID = .List_PID(Inx_Process)
        RAM_ID = .Processes_SDI(Inx_Process).RAM_ID_Process

        With .Processes_SDI(Inx_Process)
            .Priority_Level = Threads_Change_Priority(PID)
            If .Priority_Level = Low_Priority Then
                .Efficiency_Mode = True
            Else
                .Efficiency_Mode = False
            End If

            Affinity_Mask = Threads_Get_ProcessAffinity(PID, Restrict_Excel)
            If Len(Affinity_Mask) > 0 Then
                With .Process_Affinity
                    .Affinity_Mask = Split(Affinity_Mask, "_")(0)
                     Affinity_Mask = Split(Affinity_Mask, "_")(1)
                    .Active_LogicalCores = Split(Affinity_Mask, "|")(0)
                    .Count_LogicalCores = Split(Affinity_Mask, "|")(1)
                End With
            End If
        End With

        Select Case CheckProcess_State(PID)

            Case P_State_ActiveExecution
                GoSub GS¦Excel_Responsive
                If Not Flag_NotResponsive Then
                    On Error Resume Next
                    Set COM_Application = GetObject(RAM_ID): Err_Number = Err.Number
                        State_Process = COM_Application.CustomDocumentProperties("State_Process")
                        VBProject_Mode = COM_Application.VBProject.Mode
                    Set COM_Application = Nothing
                    On Error GoTo 0

                    If Err_Number <> 0 Then
                        With .Processes_SDI(Inx_Process)
                            .State_Process = P_State_NonExistent
                            .Init_Process = False
                            .RAM_ID_Process = vbNullString
                            .Handle_Process = -1: .Efficiency_Mode = False
                            .Priority_Level = Unknown_Priority
                            State_Process = P_State_NonExistent: VBProject_Mode = -1
                            With .Process_Affinity
                                .Active_LogicalCores = -1
                                .Affinity_Mask = vbNullString
                                .Count_LogicalCores = -1
                            End With
                        End With

                        Return
                    End If

                    Select Case VBProject_Mode
                        Case 0
                            If State_Process = P_State_ActiveExecution Then
                                .Processes_SDI(Inx_Process).State_Process = State_Process
                            Else
                                .Processes_SDI(Inx_Process).State_Process = P_State_UserTriggeredExecution
                            End If
                        Case 1
                            .Processes_SDI(Inx_Process).State_Process = P_State_Debugging
                        Case 2
                            If State_Process = P_State_ActiveExecution Then
                                .Processes_SDI(Inx_Process).State_Process = P_State_ManuallyStopped
                            Else
                                .Processes_SDI(Inx_Process).State_Process = State_Process
                            End If
                    End Select
                End If

            Case P_State_Terminate
                With .Processes_SDI(Inx_Process)
                    .State_Process = P_State_Terminate
                    .Init_Process = False
                    .RAM_ID_Process = vbNullString
                    .Handle_Process = -1: .Efficiency_Mode = False
                    .Priority_Level = Unknown_Priority
                    With .Process_Affinity
                        .Active_LogicalCores = -1
                        .Affinity_Mask = vbNullString
                        .Count_LogicalCores = -1
                    End With
                End With

            Case P_State_ErrorReading
                .Processes_SDI(Inx_Process).State_Process = P_State_ErrorReading

            Case P_State_NonExistent
                .Processes_SDI(Inx_Process).State_Process = P_State_NonExistent
                .List_PID(Inx_Process) = .Processes_SDI(Inx_Process).Process_IDentifier

        End Select

    End With

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````'
GS¦Excel_Responsive:

    With SDIProcess_Snapshot
        Hwnd_Process = .Processes_SDI(Inx_Process).Handle_Process

        If IsExcel_Responsive(Hwnd_Process, TIMEOUT_PROCESS) Then
            Flag_NotResponsive = False
        Else
            Flag_NotResponsive = True
            .Processes_SDI(Inx_Process).State_Process = P_State_Unresponsive
        End If
    End With

    Return
'````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------'
End Function
'=========================================================================================================='


'====================================================================================================='
Public Function Process_Get_Snapshot( _
                Optional ByVal Process_ID As Long = -1, _
                Optional ByVal dw_Reserved_1 As Long = 0& _
       ) As Process_Snapshot
'-----------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_get_snapshot-vbd_kit_interface_sdi-bas.279/
'`````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - ToDo_Parameter    | Зарезервировано без явного описания функционала
'-----------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````'
    Dim Vector_PID() As Long, Vector_Hwnd() As Long
    Dim Vector_COMApplication() As String
    Dim Vector_Responsive()     As Boolean
    '`````````````````````````````````````````````````````````````````'
    Dim SDIProcess_Snapshot As Process_Snapshot
    Dim Executable_Component As String, State_Process As Process_State
    Dim Type_Executable_Component   As MSOffice_Type_ContentFormat_SDI
    '`````````````````````````````````````````````````````````````````'
    Dim Process_Name   As String, Count_Process    As Long
    Dim Error_Message  As String, Empty_Vector()   As Variant
    Dim Excel_Process  As Object, Hwnd_XlMain      As Long
    Dim Init_Process   As Boolean, Flag_IsExcluded As Boolean
    '`````````````````````````````````````````````````````````````````'
    Dim Exclusion_Array() As Variant, PID As Long
    '`````````````````````````````````````````````````````````````````'
    Dim Dict_ProcessName As Object, Obj_WMI         As Object
    Dim Coll_Processes   As Object, Obj_VBComponent As Object
    '`````````````````````````````````````````````````````````````````'
    Dim Inx_SDIProcess As Long, Inx_VBComponent  As Long, R As Long
    Dim Inx_VBComponent_LB1 As Long, Inx_VBComponent_UB1    As Long
    Dim Inx_LB1 As Long, Inx_UB1 As Long, i As Long, N      As Long
    '`````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    Process_Get_Snapshot = SDIProcess_Snapshot
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````'
    GoSub GS¦Get_Count_Process

    ReDim Vector_PID(1 To Count_Process)
    ReDim Vector_Hwnd(1 To Count_Process)
    ReDim Vector_Responsive(1 To Count_Process)
    ReDim Vector_COMApplication(1 To Count_Process)
    '``````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    On Error Resume Next: GoSub GS¦Get_List_ExcelProcess: On Error GoTo 0
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    If Inx_SDIProcess = 0 Then Exit Function Else Inx_SDIProcess = 0

    Inx_LB1 = LBound(Vector_Hwnd, 1)
    Inx_UB1 = UBound(Vector_Hwnd, 1)

    For i = Inx_UB1 To Inx_LB1 Step -1
        If Process_ID = -1 Then
            GoSub GS¦Fill_SDIProcess_Snapshot
        Else
            If Vector_PID(i) = Process_ID Then GoSub GS¦Fill_SDIProcess_Snapshot
        End If
    Next i

    If Not Inx_SDIProcess = 0 Then
        With SDIProcess_Snapshot
            .Init_Structure = True
            .Count_Process = Inx_SDIProcess
        End With

        Call Process_Get_State(SDIProcess_Snapshot)
    End If
    '```````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Process_Get_Snapshot = SDIProcess_Snapshot: Exit Function
    '````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````'
GS¦Fill_SDIProcess_Snapshot:

    Inx_SDIProcess = Inx_SDIProcess + 1: PID = Vector_PID(i): Hwnd_XlMain = Vector_Hwnd(i)

    With SDIProcess_Snapshot
        Process_Name = Vector_COMApplication(i): Inx_VBComponent = 0

        If Vector_Responsive(i) Then
            On Error Resume Next: Set Excel_Process = GetObject(Process_Name)
            With Excel_Process
                Executable_Component = vbNullString
                Executable_Component = .CustomDocumentProperties("Executable_Component")

                If Not Len(Executable_Component) = 0 Then
                    State_Process = .CustomDocumentProperties("State_Process")
                    ReDim Additional_Component(1 To .VBProject.VBComponents.Count)

                    For Each Obj_VBComponent In .VBProject.VBComponents
                        If Not Obj_VBComponent.Type = 100 Then
                            Inx_VBComponent = Inx_VBComponent + 1
                            Additional_Component(Inx_VBComponent) = Obj_VBComponent.Name
                        End If
                    Next
                End If
            End With

            Set Excel_Process = Nothing: Init_Process = True: On Error GoTo 0

            If Len(Executable_Component) = 0 Then
                Init_Process = False
            Else
                Exclusion_Array = Array(Executable_Component, VBComponent_KernelName)

                If Inx_VBComponent > 0 Then
                    ReDim Preserve Additional_Component(1 To Inx_VBComponent)
                    Inx_VBComponent = 0
                    Inx_VBComponent_LB1 = LBound(Additional_Component, 1)
                    Inx_VBComponent_UB1 = UBound(Additional_Component, 1)

                    For R = Inx_VBComponent_LB1 To Inx_VBComponent_UB1
                        Flag_IsExcluded = False

                        For N = LBound(Exclusion_Array) To UBound(Exclusion_Array)
                            If StrComp(Additional_Component(R), _
                                       Exclusion_Array(N), vbTextCompare) = 0 Then
                                Flag_IsExcluded = True: Exit For
                            End If
                        Next N

                        If Not Flag_IsExcluded Then
                            Inx_VBComponent = Inx_VBComponent + 1
                            Additional_Component(Inx_VBComponent) = Additional_Component(R)
                        End If
                    Next R

                    If Inx_VBComponent = 0 Then
                        Additional_Component = Empty_Vector
                    Else
                        ReDim Preserve Additional_Component(1 To Inx_VBComponent)
                        Sorting_VBComponents Additional_Component, _
                                             LBound(Additional_Component), _
                                             UBound(Additional_Component)
                On Error GoTo 0
                    End If
                    Type_Executable_Component = MSDI_VBComponent_SDI
                Else
                    Type_Executable_Component = MSDI_Empty_SDI
                End If
            End If
        Else
            Executable_Component = "Unknown"
            Init_Process = True
            Additional_Component = Empty_Vector
            State_Process = P_State_NonExistent
            Type_Executable_Component = MSDI_Empty_SDI
        End If

        If Init_Process Then
            ReDim Preserve .List_PID(1 To Inx_SDIProcess)
            ReDim Preserve .Processes_SDI(1 To Inx_SDIProcess)

            With .Processes_SDI(Inx_SDIProcess)
                .Executable_Component = Executable_Component
                .Additional_Component = Additional_Component
                .Init_Process = Init_Process
                .Process_IDentifier = PID
                .Handle_Process = Hwnd_XlMain
                .RAM_ID_Process = Process_Name
                .State_Process = State_Process
                .Type_Executable_Component = Type_Executable_Component
            End With

            .List_PID(Inx_SDIProcess) = PID
        Else
            Inx_SDIProcess = Inx_SDIProcess - 1
        End If
    End With

    Return
'``````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````'
GS¦Get_List_ExcelProcess:

    Hwnd_XlMain = FindWindowEx(0, 0, "XLMAIN", vbNullString)
    Set Dict_ProcessName = CreateObject("Scripting.Dictionary")

    While Hwnd_XlMain <> 0
        Process_Name = Get_RunningProcess_ByHwnd(Hwnd_XlMain)
        If Len(Process_Name) > 0 Then
            Call GetWindowThreadProcessId(Hwnd_XlMain, PID)
            If UCase$(Get_ProcessPath_PID(PID)) Like "*\EXCEL.EXE" Then
                If Not Dict_ProcessName.Exists(Process_Name) Then
                    Inx_SDIProcess = Inx_SDIProcess + 1
                    Vector_PID(Inx_SDIProcess) = PID
                    Vector_Hwnd(Inx_SDIProcess) = Hwnd_XlMain
                    Vector_COMApplication(Inx_SDIProcess) = Process_Name

                    If IsExcel_Responsive(Hwnd_XlMain, TIMEOUT_PROCESS) Then
                        Vector_Responsive(Inx_SDIProcess) = True
                    Else
                        Vector_Responsive(Inx_SDIProcess) = False
                    End If

                    Dict_ProcessName(Process_Name) = vbNullString
                End If
            End If
        End If

        Hwnd_XlMain = FindWindowEx(0, Hwnd_XlMain, "XLMAIN", vbNullString)
    Wend

    If Not Inx_SDIProcess = 0 Then
        ReDim Preserve Vector_PID(1 To Inx_SDIProcess)
        ReDim Preserve Vector_Hwnd(1 To Inx_SDIProcess)
        ReDim Preserve Vector_Responsive(1 To Inx_SDIProcess)
        ReDim Preserve Vector_COMApplication(1 To Inx_SDIProcess)
    End If

    Return
'````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Get_Count_Process:

    Set Obj_WMI = GetObject("winmgmts:\\.\root\cimv2")
    Set Coll_Processes = Obj_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")

    Count_Process = Coll_Processes.Count: Set Obj_WMI = Nothing: Set Coll_Processes = Nothing

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------'
End Function
'====================================================================================================='

    
'==================================================================================================================='
Public Function Process_ExecuteCommand( _
                ByRef SDIProcess_Snapshot As Process_Snapshot, _
                ByVal Type_Command As Thread_TypeCommand, _
                Optional ByVal Process_ID As Long = -1, _
                Optional ByVal Thread_Parameter As String = vbNullString, _
                Optional ByVal Ignore_RunTime As Boolean = True, _
                Optional ByVal dw_Reserved_1 As Long = 0& _
       ) As Boolean
'-------------------------------------------------------------------------------------------------------------------'
' // Documentation: _
     https://www.script-coding.ru/threads/funkcija-process_executecommand-vbd_kit_interface_sdi-bas.283/
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - ToDo_Parameter | Зарезервировано без явного описания функционала
'-------------------------------------------------------------------------------------------------------------------'
    
    '```````````````````````````````````````````````````````````````````````'
    Dim Vector_Components() As Variant, Additional_Component() As Variant
    Dim Vector_Process()    As Long, Inx_Process As Long, PID  As Long
    '```````````````````````````````````````````````````````````````````````'
    Dim State_ExecuteCommand As Boolean, Controller_Result     As String
    Dim Flag_SDIError        As Boolean, Flag_ExecuteCommand   As Boolean
    Dim Flag_ControllerError As Boolean, Flag_NotResponsive    As Boolean
    '```````````````````````````````````````````````````````````````````````'
    Dim ExecuteCommand_Name  As String, Name_VBComponent       As String
    Dim Executable_Component As String, Affinity_Mask          As String
    '```````````````````````````````````````````````````````````````````````'
    Dim ExecuteCommand_Status    As Long, Hwnd_Process         As Long
    Dim ExecuteCommand_ErrorCode As Long, Count_VBComponent    As Long
    Dim Module_Header_Delimiter  As String, Error_Message      As String
    Dim Module_Delimiter         As String, Full_Name          As String
    '```````````````````````````````````````````````````````````````````````'
    Dim Dict_AdditionalComponent As Object, Count_PID As Long, i As Long
    Dim Type_Data As MSOffice_Type_ContentFormat_SDI, RAM_ID     As String
    '```````````````````````````````````````````````````````````````````````'
    Dim Excel_Process As Object, Inx_LB1 As Long, Inx_UB1 As Long, N As Long
    '```````````````````````````````````````````````````````````````````````'

    '`````````````````````````````'
    Process_ExecuteCommand = False
    '`````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI: Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function

    If Not Glb_MSOffice_Type_Application = Type_MSOffice_Excel Then
        Error_Message = "Поддержка интерфейса SDI и API реализован только для приложения MS Excel!"
        Show_ErrorMessage_Immediate Error_Message, "Невозможно эксплуатировать интерфейс SDI!"
        Exit Function
    End If
    '``````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````'
    Count_PID = SDIProcess_Snapshot.Count_Process: If Count_PID = 0 Then Exit Function
    '`````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'
    If Process_ID = -1 Then
        ReDim Vector_Process(1 To Count_PID)
        For i = 1 To Count_PID: Vector_Process(i) = i: Next i
    Else
        Inx_Process = -1

        For i = 1 To Count_PID
            If Process_ID = SDIProcess_Snapshot.List_PID(i) Then
                Inx_Process = i: Exit For
            End If
        Next i

        If Inx_Process = -1 Then
            Error_Message = "Невозможно передать команду процессу, т.к. процесс отсутствует в структуре (" _
                                                                                        & Process_ID & ")"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
            Exit Function
        End If

        Vector_Process = Array(Inx_Process)
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````'
    Inx_LB1 = LBound(Vector_Process, 1)
    Inx_UB1 = UBound(Vector_Process, 1)
    State_ExecuteCommand = True
    '``````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Type_Command
        Case CommandThread_Initialize:                ExecuteCommand_Name = "Command_Initialize"
        Case CommandThread_Terminate:                 ExecuteCommand_Name = "Command_Terminate"
        Case CommandThread_Run:                       ExecuteCommand_Name = "Command_Run"
        Case CommandThread_Halt:                      ExecuteCommand_Name = "Command_Halt"
        Case CommandThread_Save:                      ExecuteCommand_Name = "Command_Save"
        Case CommandThread_Resume:                    ExecuteCommand_Name = "Command_Resume"
        Case CommandThread_Suspended:                 ExecuteCommand_Name = "Command_Suspended"
        Case CommandThread_AbortExecution:            ExecuteCommand_Name = "Command_AbortExecution"
        Case CommandThread_ChangePriority:            ExecuteCommand_Name = "Command_ChangePriority"
        Case CommandThread_EfficiencyMode:            ExecuteCommand_Name = "Command_EfficiencyMode"
        Case CommandThread_ProcessAffinity:           ExecuteCommand_Name = "Command_ProcessAffinity"
        Case CommandThread_VBProject_AddComponent:    ExecuteCommand_Name = "Command_AddComponent"
        Case CommandThread_VBProject_RemoveComponent: ExecuteCommand_Name = "Command_RemoveComponent"
        Case CommandThread_VBProject_UpdateComponent: ExecuteCommand_Name = "Command_UpdateComponent"

        Case Else: ExecuteCommand_Name = "Command_Unknown"
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    With SDIProcess_Snapshot

        Module_Delimiter = Chr$(95) & Chr$(47) & "_" & Chr$(92) & Chr$(95)
        Module_Header_Delimiter = Chr$(95) & Chr$(92) & "_" & Chr$(47) & Chr$(95)

        For i = Inx_LB1 To Inx_UB1
        
            Inx_Process = Vector_Process(i): PID = .List_PID(Inx_Process)
            Hwnd_Process = .Processes_SDI(Inx_Process).Handle_Process
            RAM_ID = .Processes_SDI(Inx_Process).RAM_ID_Process

            Select Case CheckProcess_State(PID)
                Case P_State_ErrorReading
                    .Processes_SDI(Inx_Process).State_Process = P_State_ErrorReading
                    Error_Message = "Невозможно передать команду """ & ExecuteCommand_Name & _
                                    """ процессу! Ошибка чтения процесса (" & PID & ")"
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
                    State_ExecuteCommand = False: GoTo GT¦Next_SDIProcess

                Case P_State_Terminate, P_State_NonExistent
                    .Processes_SDI(Inx_Process).State_Process = P_State_Terminate
                    Error_Message = "Невозможно передать команду """ & ExecuteCommand_Name & _
                                    """ процессу, т.к. процесс не существует (" & PID & ")"
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
                    State_ExecuteCommand = False: GoTo GT¦Next_SDIProcess
            End Select

            Select Case Type_Command
                Case CommandThread_Resume:          GoSub GS¦ExecuteCommand_Resume:          GoTo GT¦Next_SDIProcess
                Case CommandThread_Suspended:       GoSub GS¦ExecuteCommand_Suspended:       GoTo GT¦Next_SDIProcess
                Case CommandThread_ChangePriority:  GoSub GS¦ExecuteCommand_ChangePriority:  GoTo GT¦Next_SDIProcess
                Case CommandThread_EfficiencyMode:  GoSub GS¦ExecuteCommand_EfficiencyMode:  GoTo GT¦Next_SDIProcess
                Case CommandThread_ProcessAffinity: GoSub GS¦ExecuteCommand_ProcessAffinity: GoTo GT¦Next_SDIProcess
            End Select
    
            GoSub GS¦Excel_Responsive

            If Flag_NotResponsive Then
                If Not Type_Command = CommandThread_Terminate Then
                    .Processes_SDI(Inx_Process).State_Process = P_State_Unresponsive
                    Error_Message = "Невозможно передать команду """ & ExecuteCommand_Name & _
                                    """ процессу! Процесс не отвечает (" & PID & ")"
                    Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
                    State_ExecuteCommand = False: GoTo GT¦Next_SDIProcess
                End If
            End If

            Select Case Type_Command
                Case CommandThread_Initialize:                GoSub GS¦ExecuteCommand_Initialize
                Case CommandThread_Halt:                      GoSub GS¦ExecuteCommand_Halt
                Case CommandThread_AbortExecution:            GoSub GS¦ExecuteCommand_AbortExecution
                Case CommandThread_Terminate:                 GoSub GS¦ExecuteCommand_Terminate
                Case CommandThread_Run:                       GoSub GS¦ExecuteCommand_Run
                Case CommandThread_Save:                      GoSub GS¦ExecuteCommand_Save
                Case CommandThread_VBProject_AddComponent:    GoSub GS¦ExecuteCommand_VBProject_AddComponent
                Case CommandThread_VBProject_UpdateComponent: GoSub GS¦ExecuteCommand_VBProject_UpdateComponent
                Case CommandThread_VBProject_RemoveComponent: GoSub GS¦ExecuteCommand_VBProject_RemoveComponent
            End Select

GT¦Next_SDIProcess:
        Next i

    End With

    Process_ExecuteCommand = State_ExecuteCommand

    Exit Function
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Controller_SendCommand:

    On Error Resume Next: Set Excel_Process = GetObject(RAM_ID)
    If Excel_Process Is Nothing Then GoSub GS¦Process_IsUnresponsive: Return

    Controller_Result = Excel_Process _
                           .Application.Run(VBComponent_KernelName & ".Controller_Thread", _
                                            CLng(Type_Command), Thread_Parameter _
                        )

    Set Excel_Process = Nothing: On Error GoTo 0

    If InStr(1, Controller_Result, "_") > 0 Then
        ExecuteCommand_Status = CLng(Split(Controller_Result, "_")(1))
        ExecuteCommand_ErrorCode = CLng(Split(Controller_Result, "_")(0))
    Else
        GoSub GS¦Structure_Process_IsBroken: State_ExecuteCommand = False: Return
    End If

    Flag_ControllerError = False: Flag_SDIError = False
    Flag_ExecuteCommand = CBool(ExecuteCommand_ErrorCode)

    If Flag_ExecuteCommand Then
        Flag_ControllerError = True
        GoSub GS¦Process_ControllerError: On Error GoTo 0: Return
    End If

    Select Case ExecuteCommand_Status
        Case StatusCommandThread_Done
            Select Case Err.Number
                Case 1004&: GoSub GS¦Structure_Process_IsBroken: Flag_SDIError = True
                Case 57097: GoSub GS¦Process_IsUnavailable:      Flag_SDIError = True
            End Select

        Case StatusCommandThread_NotFound
            Error_Message = "Команда """ & ExecuteCommand_Name & _
                            """ не найдена в контроллере процесса - " & _
                            "[" & RAM_ID & "(PID = " & PID & ")]"
            Show_ErrorMessage_Immediate Error_Message, "Внутренняя ошибка SDI_API"
            Flag_ControllerError = True: State_ExecuteCommand = False

        Case StatusCommandThread_NeedAbort
            Error_Message = "Команда """ & ExecuteCommand_Name & _
                            """ не может быть выполнена при работающем коде в процессе - " & _
                            "[" & RAM_ID & "(PID = " & PID & ")] " & _
                            "и параметре ""Ignore_RunTime""=False"
            Show_ErrorMessage_Immediate Error_Message, "Исключение SDI_API"
            Flag_ControllerError = True: State_ExecuteCommand = False
    End Select

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Initialize:

    GoSub GS¦Controller_SendCommand

    If Not Flag_ExecuteCommand Then
        If Not Flag_SDIError Then
            With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
                .State_Process = P_State_ActiveExecution
            End With
        End If
    End If

    Return
'``````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Halt:

    GoSub GS¦Controller_SendCommand
    Flag_SDIError = Not Process_Get_State(SDIProcess_Snapshot, PID)

    With SDIProcess_Snapshot.Processes_SDI(i)
        If .State_Process = P_State_ManuallyStopped Then
            Error_Message = "Не удаётся приостановить выполнение кода в процессе - " & _
                            "[" & RAM_ID & "(PID = " & PID & ")] " & _
                            "Возможно запущенный код не обрабатывает данное событие!"
            Show_ErrorMessage_Immediate Error_Message, "Внутренняя ошибка шаблона SDI_API"

            .State_Process = P_State_ActiveExecution: State_ExecuteCommand = False
        End If
    End With

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_AbortExecution:

    On Error Resume Next: Set Excel_Process = GetObject(RAM_ID)
    If Excel_Process Is Nothing Then GoSub GS¦Process_IsUnresponsive: Return

    Controller_Result = Excel_Process _
                           .Application.Run(VBComponent_KernelName & ".Controller_Thread", _
                                            CLng(Type_Command), Thread_Parameter _
                        )

    Set Excel_Process = Nothing: On Error GoTo 0

    Flag_SDIError = Not Process_Get_State(SDIProcess_Snapshot, PID)

    If Flag_SDIError Then
        Error_Message = "Не удаётся получить состояние процесса после приостановки - " & _
                        "[" & RAM_ID & "(PID = " & PID & ")]"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка получения состояния"
        State_ExecuteCommand = False
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Terminate:

    If Flag_NotResponsive Then
        Flag_ExecuteCommand = KillProcess_PID(PID)
    Else
        On Error Resume Next: Set Excel_Process = GetObject(RAM_ID)
        If Excel_Process Is Nothing Then GoSub GS¦Process_IsUnresponsive: Return

        Excel_Process.Application.Run VBComponent_KernelName & ".Controller_Thread", CLng(Type_Command)
        Set Excel_Process = Nothing: Full_Name = GetObject(RAM_ID).codeName: On Error GoTo 0

        If Len(Full_Name) = 0 Then
            Flag_ExecuteCommand = True
        Else
            Flag_ExecuteCommand = False
        End If
    End If

    If Not Flag_ExecuteCommand Then Flag_ExecuteCommand = KillProcess_PID(PID)

    If Flag_ExecuteCommand Then
        With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
            .Handle_Process = -1
            .Init_Process = False
            .Process_IDentifier = -1
            .State_Process = P_State_Terminate
            .Efficiency_Mode = False
            .Priority_Level = Unknown_Priority

            With .Process_Affinity
                .Active_LogicalCores = -1
                .Affinity_Mask = vbNullString
                .Count_LogicalCores = -1
            End With
        End With
    Else
        Error_Message = "Не удаётся закрыть процесс! " & _
                        "Проверьте изолированный процесс - " & RAM_ID & "(PID = " & PID & ")"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка подключения к процессу Excel"
        State_ExecuteCommand = False
    End If

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Run:

    GoSub GS¦Controller_SendCommand

    If Not Flag_ExecuteCommand Then
        If Not Flag_SDIError Then
            With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
                .State_Process = P_State_ActiveExecution
            End With
        End If
    End If

    Return
'``````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Save:

    If Len(Thread_Parameter) = 0 Then
        Error_Message = "Для сохранения процесса, требуется указать папку, куда сохранить файл!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    If Not Right$(Thread_Parameter, 1) = "\" Then Thread_Parameter = Thread_Parameter & "\"

    If Get_SemanticDataType(Thread_Parameter) = MSDI_Folder_SDI Then
        Full_Name = Thread_Parameter & RAM_ID & ".xlsx"

        If File_IsOpen(Full_Name) Then
            Error_Message = "Процесс Excel пытается перезаписать открытый файл. " & _
                            "Закройте файл (" & Full_Name & ") и повторите попытку"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка сохранения процесса Excel"
            State_ExecuteCommand = False: Return
        End If

        On Error Resume Next: Set Excel_Process = GetObject(RAM_ID)
        If Excel_Process Is Nothing Then GoSub GS¦Process_IsUnresponsive: Return

        Excel_Process.SaveCopyAs Full_Name: Set Excel_Process = Nothing
        If CBool(Err.Number) Then State_ExecuteCommand = False

        On Error GoTo 0
    Else
        Error_Message = "Переданный путь не является директорией (папкой) " & _
                        "или данная директория не существует - " & Thread_Parameter
        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_VBProject_AddComponent:
GS¦ExecuteCommand_VBProject_UpdateComponent:

    If Len(Thread_Parameter) = 0 Then
        Error_Message = "Параметр с перечнем модулей для добавления или обновления - пуст! " & _
                        "Заполните параметр ""Thread_Parameter"" соответствующими именами модулей!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    Vector_Components = Application.Trim(Split(Thread_Parameter, "|"))
    Inx_LB1 = LBound(Vector_Components, 1)
    Inx_UB1 = UBound(Vector_Components, 1)
    Count_VBComponent = Inx_LB1 - 1

    For N = Inx_LB1 To Inx_UB1
        Name_VBComponent = Vector_Components(N)
        If ExecuteCommand_Name = "Command_AddComponent" Then
            If Name_VBComponent = VBComponent_KernelName Then
                Error_Message = "Нельзя добавить компонент с именем модуля ядра процесса - """ & _
                                 VBComponent_KernelName & """ " & " (только через модификацию API-кода)!"
                Show_ErrorMessage_Immediate Error_Message, "Операция недоступна на уровне SDI_API"
                Exit Function
            End If
        End If

        If Not Len(Name_VBComponent) = 0 Then
            Count_VBComponent = Count_VBComponent + 1
            Type_Data = Get_SemanticDataType(Name_VBComponent)
            If Type_Data = MSDI_VBComponent_SDI Then
                Vector_Components(N) = Name_VBComponent & Module_Delimiter & Get_ModuleCode(Name_VBComponent)
            ElseIf Type_Data = MSDI_File_SDI Then
            Else
                Error_Message = "Компонент """ & Name_VBComponent & """ " & _
                                "Не найден ни в файловой системе, ни в текущем проекте VBA!"
                Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
                Exit Function
            End If
        End If
    Next N

    If Count_VBComponent = Inx_LB1 - 1 Then
        If ExecuteCommand_Name = "Command_AddComponent" Then
            Error_Message = "Не найдены компоненты, которые можно добавить! Проверьте параметр " & _
                        """Thread_Parameter"" на корректность заполнения! - " & Thread_Parameter
        Else
            Error_Message = "Не найдены компоненты, которые можно обновить! Проверьте параметр " & _
                        """Thread_Parameter"" на корректность заполнения! - " & Thread_Parameter
        End If

        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    If Count_VBComponent <> Inx_UB1 Then ReDim Preserve Vector_Components(Inx_LB1 To Count_VBComponent)
    Thread_Parameter = Join(Vector_Components, Module_Header_Delimiter)

    GoSub GS¦Controller_SendCommand: GoSub GS¦Update_VBComponentList

    Return
'````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_VBProject_RemoveComponent:

    If Len(Thread_Parameter) = 0 Then
        Error_Message = "Параметр с перечнем модулей для удаления - пуст! " & _
                        "Заполните параметр ""Thread_Parameter"" соответствующими именами модулей!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    Vector_Components = Application.Trim(Split(Thread_Parameter, "|"))
    Inx_LB1 = LBound(Vector_Components, 1)
    Inx_UB1 = UBound(Vector_Components, 1)
    Count_VBComponent = Inx_LB1 - 1

    For N = Inx_LB1 To Inx_UB1
        Name_VBComponent = Vector_Components(N)
        If Not Len(Name_VBComponent) = 0 Then
            Count_VBComponent = Count_VBComponent + 1
            If Name_VBComponent = VBComponent_KernelName Then
                Error_Message = "Нельзя удалить компонент с контроллером исполняемого потока """ & _
                                 Name_VBComponent & """ " & " (только через модификацию API-кода)!"
                Show_ErrorMessage_Immediate Error_Message, "Операция недоступна на уровне SDI_API"
                Exit Function
            End If
        End If
    Next N

    If Count_VBComponent = Inx_LB1 - 1 Then
        Error_Message = "Не один указанный компоненты не может быть удален! Проверьте параметр " & _
                        """Thread_Parameter"" на корректность заполнения! - " & Thread_Parameter
        Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
        Exit Function
    End If

    If Count_VBComponent <> Inx_UB1 Then ReDim Preserve Vector_Components(Inx_LB1 To Count_VBComponent)
    Thread_Parameter = Join(Vector_Components, "|")

    GoSub GS¦Controller_SendCommand: GoSub GS¦Update_VBComponentList

    Return
'``````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Resume:

    Flag_ExecuteCommand = Threads_Change_State(PID, T_State_Resume, Restrict_Excel)

    If Flag_ExecuteCommand Then
        Flag_ExecuteCommand = Process_Get_State(SDIProcess_Snapshot, PID)
        If Not Flag_ExecuteCommand Then
            Error_Message = "Не удаётся получить состояние процесса после возобновления - " & _
                            "[" & RAM_ID & "(PID = " & PID & ")]"
            Show_ErrorMessage_Immediate Error_Message, "Ошибка получения состояния"
        End If
    Else
        Error_Message = "Не удаётся возобновить состояние процесса - " & _
                        "[" & RAM_ID & "(PID = " & PID & ")]"
        Show_ErrorMessage_Immediate Error_Message, "Системная ошибка WinAPI"
        State_ExecuteCommand = False
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_Suspended:

    Flag_ExecuteCommand = Threads_Change_State(PID, T_State_Suspended, Restrict_Excel)
    If Not Flag_ExecuteCommand Then
        Error_Message = "Не удаётся приостановить состояние процесса - " & _
                        "[" & RAM_ID & "(PID = " & PID & ")]"
        Show_ErrorMessage_Immediate Error_Message, "Системная ошибка WinAPI"
        State_ExecuteCommand = False
    Else
        SDIProcess_Snapshot.Processes_SDI(Inx_Process).State_Process = P_State_Suspended
    End If

    Return
'````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_ChangePriority:

    SDIProcess_Snapshot.Processes_SDI(Inx_Process).Priority_Level = _
    Threads_Change_Priority(PID, CLng(Thread_Parameter), Restrict_Excel)

    Return
'```````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_EfficiencyMode:

    With SDIProcess_Snapshot
        If Threads_EfficiencyMode(PID, CBool(Thread_Parameter), Restrict_Excel) Then
            With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
                .Efficiency_Mode = CBool(Thread_Parameter)
            End With
        End If

        SDIProcess_Snapshot. _
        Processes_SDI(Inx_Process).Priority_Level = Threads_Change_Priority(PID)

        If .Processes_SDI(Inx_Process).Efficiency_Mode <> CBool(Thread_Parameter) Then
            Error_Message = "Не удаётся задать режим эффективности для процесса - " & _
                            "[" & RAM_ID & "(PID = " & PID & ")]"
            Show_ErrorMessage_Immediate Error_Message, "Системная ошибка WinAPI"
            State_ExecuteCommand = False
        End If
    End With

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````'
GS¦ExecuteCommand_ProcessAffinity:

    Flag_ExecuteCommand = _
    Threads_Set_ProcessAffinity(PID, CStr(Thread_Parameter), Restrict_Excel)

    If Not Flag_ExecuteCommand Then
        Error_Message = "Не удаётся задать маску сходства для процесса - " & _
                        "[" & RAM_ID & "(PID = " & PID & ")]"
        Show_ErrorMessage_Immediate Error_Message, "Системная ошибка WinAPI"
        State_ExecuteCommand = False
    End If

    Affinity_Mask = Threads_Get_ProcessAffinity(PID, Restrict_Excel)
    If Len(Affinity_Mask) > 0 Then
        With SDIProcess_Snapshot.Processes_SDI(Inx_Process).Process_Affinity
            .Affinity_Mask = Split(Affinity_Mask, "_")(0)
             Affinity_Mask = Split(Affinity_Mask, "_")(1)
            .Active_LogicalCores = Split(Affinity_Mask, "|")(0)
            .Count_LogicalCores = Split(Affinity_Mask, "|")(1)
        End With
    Else
        Error_Message = "Не удаётся получить маску сходства после изменения - " & _
                        "[" & RAM_ID & "(PID = " & PID & ")]"
        Show_ErrorMessage_Immediate Error_Message, "Системная ошибка WinAPI"
        State_ExecuteCommand = False

        With SDIProcess_Snapshot.Processes_SDI(Inx_Process).Process_Affinity
            .Affinity_Mask = vbNullString
            .Active_LogicalCores = -1
            .Count_LogicalCores = -1
        End With
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````'
GS¦Update_VBComponentList:

    Select Case Type_Command
        Case CommandThread_VBProject_AddComponent, _
             CommandThread_VBProject_RemoveComponent
        Case Else: Return
    End Select

    If Flag_ExecuteCommand Or Flag_SDIError Then Return

    With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
        Executable_Component = .Executable_Component
        Set Dict_AdditionalComponent = CreateObject("Scripting.Dictionary")

        On Error Resume Next
        Inx_LB1 = LBound(.Additional_Component, 1)
        If Err.Number = 0 Then
            On Error GoTo 0: Additional_Component = .Additional_Component
            Inx_LB1 = LBound(Additional_Component, 1)
            Inx_UB1 = UBound(Additional_Component, 1)
            For N = Inx_LB1 To Inx_UB1
                Dict_AdditionalComponent(Additional_Component(N)) = vbNullString
            Next N
        Else
            ReDim Additional_Component(1 To 1)
        End If
        On Error GoTo 0
    End With

    Inx_LB1 = LBound(Vector_Components, 1)
    Inx_UB1 = UBound(Vector_Components, 1)

    Select Case Type_Command
        Case CommandThread_VBProject_AddComponent
            For N = Inx_LB1 To Inx_UB1
                Name_VBComponent = Split(Vector_Components(N), Module_Delimiter)(0)
                If Right(Name_VBComponent, 4) = ".bas" Then
                    Name_VBComponent = Mid$(Name_VBComponent, _
                                    InStrRev(Name_VBComponent, "\") + 1)
                    Name_VBComponent = Mid$(Name_VBComponent, 1, _
                                    InStrRev(Name_VBComponent, ".") - 1)
                End If

                Dict_AdditionalComponent(Name_VBComponent) = vbNullString
            Next N

        Case CommandThread_VBProject_RemoveComponent
            For N = Inx_LB1 To Inx_UB1
                Name_VBComponent = Vector_Components(N)
                If Name_VBComponent = Executable_Component Then
                    Executable_Component = vbNullString
                End If

                If Dict_AdditionalComponent.Exists(Name_VBComponent) Then
                    Dict_AdditionalComponent.Remove Name_VBComponent
                End If
            Next N
    End Select

    If Dict_AdditionalComponent.Count = 0 Then
        Erase Additional_Component
    Else
        Additional_Component = Application.Trim$(Dict_AdditionalComponent.keys)
        Sorting_VBComponents Additional_Component, LBound(Additional_Component), _
                                                   UBound(Additional_Component)
    End If

    With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
        .Executable_Component = Executable_Component
        .Additional_Component = Additional_Component
    End With

    Return
'```````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````'
GS¦Excel_Responsive:

    If Not IsExcel_Responsive(Hwnd_Process, TIMEOUT_PROCESS) Then
        Flag_NotResponsive = True
    Else
        Flag_NotResponsive = False
    End If

    Return
'`````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GS¦Process_ControllerError:

    Error_Message = "Внутренняя ошибка при попытке выполнить команду """ & _
                     ExecuteCommand_Name & """ (" & ExecuteCommand_ErrorCode & ")! " & _
                    "Проверьте изолированный процесс - " & RAM_ID & "(PID = " & PID & ")"
    Show_ErrorMessage_Immediate Error_Message, "Ошибка в контроллере процесса Excel"
    State_ExecuteCommand = False

    Return
'`````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GS¦Process_IsUnresponsive:

    Error_Message = "Не удаётся подключится к процессу через COM вызов! " & _
                    "Проверьте изолированный процесс - " & RAM_ID & "(PID = " & PID & ")"
    Show_ErrorMessage_Immediate Error_Message, "Ошибка подключения к процессу Excel"
    State_ExecuteCommand = False

    Return
'`````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````````````````'
GS¦Process_IsUnavailable:

    Process_Get_State SDIProcess_Snapshot, SDIProcess_Snapshot.List_PID(Inx_Process)

    With SDIProcess_Snapshot.Processes_SDI(Inx_Process)
        If .State_Process = P_State_Debugging Then
            Error_Message = "Процесс - " & PID & " не может получить команду """ & _
                            ExecuteCommand_Name & """! " & _
                            "Данный процесс находится в режиме отладки!"
        Else
            Error_Message = "Процесс - " & PID & " не может получить команду """ & _
                            ExecuteCommand_Name & """! " & _
                            "Проверьте процесс в ручном режиме!"
        End If
    End With

    Show_ErrorMessage_Immediate Error_Message, "Ошибка передачи команды процессу Excel"
    State_ExecuteCommand = False

    Return
'``````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Structure_Process_IsBroken:

    Error_Message = "Изменена структура ядра процесса или вовсе отсутствует модуль ядра. " & _
                    "Проверьте изолированный процесс - " & RAM_ID & "(PID = " & PID & ")"
    Show_ErrorMessage_Immediate Error_Message, "Нарушена структура ядра изолированного процесса"
    State_ExecuteCommand = False

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------'
End Function
'==================================================================================================================='


'======================================================================================='
Public Function Threads_Change_State( _
                ByVal Process_Metric As Variant, _
                ByVal State_Thread As Thread_State, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As Boolean
'---------------------------------------------------------------------------------------'
' // Documentation: In_Process
'---------------------------------------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim Handle_Snapshot As Long
    Dim Handle_Thread   As Long
    Dim ThreadEntry32   As WinAPI_ThreadEntry32
    Dim Processes_PID   As Variant
    Dim Error_Message   As String
    '``````````````````````````````````````````'
    Dim result  As Boolean, Inx_UB1 As Long
    Dim Process_ID As Long, i       As Long
    '``````````````````````````````````````````'

    '``````````````````````````'
    Threads_Change_State = False
    '``````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    If Not IsNumeric(Process_Metric) Then
        Processes_PID = Get_RunningProcesses_ByName(Process_Metric)
        Inx_UB1 = UBound(Processes_PID, 1)
        If Inx_UB1 = -1 Then
            Error_Message = "Не найдено ни одного процесса по имени: " & Process_Metric
            Show_ErrorMessage_Immediate Error_Message, "Нет активных процессов!"
            Exit Function
        End If
    Else
        ReDim Processes_PID(1 To 1) As Long
        Inx_UB1 = 1: Processes_PID(1) = Process_Metric
    End If
    '``````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````'
    Handle_Snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0&)
    '````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````'
    If Handle_Snapshot <> INVALID_HANDLE_VALUE Then
        ThreadEntry32.dwSize = Len(ThreadEntry32)
        result = Thread32First(Handle_Snapshot, ThreadEntry32)
        Do While result
            For i = 1 To Inx_UB1
                Process_ID = Processes_PID(i)
                If Process_Restrict = Restrict_Excel Then
                    If UCase$(Get_ProcessPath_PID(Process_ID)) Like "*EXCEL.EXE" Then
                        GoSub GS¦T_ChangeState
                    End If
                Else
                    GoSub GS¦T_ChangeState
                End If
            Next i
            result = Thread32Next(Handle_Snapshot, ThreadEntry32)
        Loop
        CloseHandle Handle_Snapshot
    End If

    Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````'
GS¦T_ChangeState:

    If ThreadEntry32.rh32OwnerProcessID = Process_ID Then
         Handle_Thread = OpenThread(NT_THREAD_SUSPEND_RESUME, False, _
                                    ThreadEntry32.th32ThreadID)
         If Handle_Thread <> 0 Then
             Select Case State_Thread
                 Case T_State_Resume:     ResumeThread Handle_Thread
                 Case T_State_Suspended: SuspendThread Handle_Thread
             End Select
             Threads_Change_State = True:   CloseHandle Handle_Thread
         End If
    End If

    Return
'``````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------'
End Function
'======================================================================================='


'================================================================================================='
Public Function Threads_Change_Priority( _
                ByVal Process_ID As Long, _
                Optional ByVal Priority_Level As Process_Priority = Unknown_Priority, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As Process_Priority
'-------------------------------------------------------------------------------------------------'
' // Documentation: In_Process
'-------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Handle_Process As LongPtr
    Dim Error_Message  As String, Change_Result As Long
    '``````````````````````````````````````````````````'

    '````````````````````````````````````````'
    Threads_Change_Priority = Unknown_Priority
    '````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If Process_Restrict = Restrict_Excel Then
        If Not UCase$(Get_ProcessPath_PID(Process_ID)) Like "*EXCEL.EXE" Then
            Exit Function
        End If
    End If
    '`````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If Priority_Level = Unknown_Priority Then
        Threads_Change_Priority = Get_ProcessPriority_PID(Process_ID)
    Else
        Handle_Process = OpenProcess(NT_PROCESS_SET_INFORMATION, 0&, Process_ID)

        If CBool(Handle_Process) Then
            Change_Result = SetPriorityClass(Handle_Process, Priority_Level)
            CloseHandle Handle_Process
            Threads_Change_Priority = Get_ProcessPriority_PID(Process_ID)
        Else
            Error_Message = "Ошибка при попытке получить доступ к процессу: " & Err.LastDllError
            Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
        End If
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'================================================================================================='
Public Function Threads_Get_ProcessAffinity( _
                ByVal Process_ID As Long, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As String
'-------------------------------------------------------------------------------------------------'
' // Documentation: In_Process
'-------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````'
    Dim Handle_Process  As LongPtr
    Dim Process_AffMask As LongPtr
    Dim System_AffMask  As LongPtr
    '`````````````````````````````````````'
    Dim Sys_Info As WinAPI_SystemInfo
    Dim Active_LogicalСores As Long
    Dim Count_LogicalСores  As Long
    '`````````````````````````````````````'
    Dim Bit_Mask      As LongPtr
    Dim Error_Message As String
    Dim Binary_String As String, i As Long
    '`````````````````````````````````````'

    '````````````````````````````````````````'
    Threads_Get_ProcessAffinity = vbNullString
    '````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If Process_Restrict = Restrict_Excel Then
        If Not UCase$(Get_ProcessPath_PID(Process_ID)) Like "*EXCEL.EXE" Then
            Exit Function
        End If
    End If
    '`````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    GetSystemInfo Sys_Info: Count_LogicalСores = Sys_Info.dwNumberOfProcessors
    '`````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````'
    Handle_Process = OpenProcess(NT_PROCESS_QUERY_INFORMATION, 0&, Process_ID)

    If Handle_Process = 0 Then
        Error_Message = "Ошибка при попытке получить доступ к процессу: " & Err.LastDllError
        Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
        Exit Function
    End If
    '````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If GetProcessAffinityMask(Handle_Process, Process_AffMask, System_AffMask) = 0 Then
        Error_Message = "Ошибка при попытке получить маску сходства процесса: " & Err.LastDllError
        Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
    Else
        GoSub GS¦AffinityMask_ToBinaryString
        Threads_Get_ProcessAffinity = Binary_String & "_" & _
                                     Active_LogicalСores & "|" & Count_LogicalСores
    End If

    CloseHandle Handle_Process: Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````'
GS¦AffinityMask_ToBinaryString:

    Bit_Mask = 1
    For i = 0 To Count_LogicalСores - 1
        If (Process_AffMask And Bit_Mask) > 0 Then
            Binary_String = Binary_String & "1"
            Active_LogicalСores = Active_LogicalСores + 1
        Else
            Binary_String = Binary_String & "0"
        End If
        Bit_Mask = Bit_Mask * 2
    Next i

    Return
'`````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'================================================================================================='
Public Function Threads_Set_ProcessAffinity( _
                ByVal Process_ID As Long, _
                ByRef Process_Mask As String, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As Boolean
'-------------------------------------------------------------------------------------------------'
' // Documentation: In_Process
'-------------------------------------------------------------------------------------------------'

    '`````````````````````````````````````'
    Dim Handle_Process  As LongPtr
    Dim Affinity_Mask   As LongPtr
    Dim Len_ProcessMask As Long, i As Long
    Dim Error_Message   As String
    '`````````````````````````````````````'
    #If x64_Soft Then
        Const Max_Bits As Long = 64
    #Else
        Const Max_Bits As Long = 32
    #End If
    '`````````````````````````````````````'

    '``````````````````````````````````'
    Threads_Set_ProcessAffinity = False
    '``````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````'
    If Process_Restrict = Restrict_Excel Then
        If Not UCase$(Get_ProcessPath_PID(Process_ID)) Like "*EXCEL.EXE" Then
            Exit Function
        End If
    End If
    '````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    Handle_Process = OpenProcess(NT_PROCESS_SET_INFORMATION, 0, Process_ID)

    If Handle_Process = 0 Then
        Error_Message = "Ошибка при попытке получить доступ к процессу: " & Err.LastDllError
        Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
        Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````'
    GoSub GS¦BinaryString_ToAffinityMask

    Select Case Affinity_Mask
        Case -1:
            Error_Message = "Длина маски сходства превышает допустимую: " & Max_Bits
            Show_ErrorMessage_Immediate _
                 Error_Message, "Некорректная маска для сходства процесса: " & Process_Mask
            CloseHandle Handle_Process: Exit Function
        Case 0:
            Error_Message = "Маска должна содержать хотя-бы один рабочий поток процессора!"
            Show_ErrorMessage_Immediate _
                 Error_Message, "Некорректная маска для сходства процесса: " & Process_Mask
            CloseHandle Handle_Process: Exit Function
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````'
    If SetProcessAffinityMask(Handle_Process, Affinity_Mask) = 0 Then
        Error_Message = "Ошибка при попытке изменить сходства процесса: " & Err.LastDllError & _
                        ". Возможно длина маски превышает количество логических ядер процессора!"
        Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
    Else
        Threads_Set_ProcessAffinity = True
    End If

    CloseHandle Handle_Process: Exit Function
    '`````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````'
GS¦BinaryString_ToAffinityMask:

    Affinity_Mask = 0: Len_ProcessMask = Len(Process_Mask)
    If Len_ProcessMask > Max_Bits Then Affinity_Mask = -1: Exit Function

    For i = Len_ProcessMask To 1 Step -1
        Affinity_Mask = Affinity_Mask * 2
        If Mid$(Process_Mask, i, 1) = "1" Then Affinity_Mask = Affinity_Mask + 1
    Next i

    Return
'```````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------'
End Function
'================================================================================================='


'====================================================================================================='
Public Function Threads_EfficiencyMode( _
                ByVal Process_Metric As Variant, _
                Optional ByVal Activate_EfficiencyMode As Boolean = True, _
                Optional ByRef Process_Restrict As Process_RestrictType = Restrict_No _
       ) As Boolean
'-----------------------------------------------------------------------------------------------------'
' // Documentation: In_Process
'-----------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````'
    Dim Handle_Process   As LongPtr, Processes_PID    As Variant
    Dim Result_Priority  As Long, Result_EfficiencyMode  As Long
    Dim Power_Throttling As WinAPI_Process_PowerThrottling_State
    '```````````````````````````````````````````````````````````'
    Dim Error_Message As String, Process_ID As Long
    Dim Final_Result  As Boolean, Inx_UB1   As Long, i As Long
    '```````````````````````````````````````````````````````````'
    Const Process_PowerThrottling As Long = 4&
    '```````````````````````````````````````````````````````````'

    '`````````````````````````````'
    Threads_EfficiencyMode = False
    '`````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    If Not IsNumeric(Process_Metric) Then
        Processes_PID = Get_RunningProcesses_ByName(Process_Metric)
        Inx_UB1 = UBound(Processes_PID, 1)
        If Inx_UB1 = -1 Then
            Error_Message = "Не найдено ни одного процесса по имени: " & Process_Metric
            Show_ErrorMessage_Immediate Error_Message, "Нет активных процессов!"
            Exit Function
        End If
    Else
        ReDim Processes_PID(1 To 1) As Long
        Inx_UB1 = 1: Processes_PID(1) = Process_Metric
    End If

    Final_Result = True
    '``````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = 1 To Inx_UB1
        Process_ID = Processes_PID(i)
        Handle_Process = OpenProcess( _
                         NT_PROCESS_SET_INFORMATION Or NT_PROCESS_QUERY_INFORMATION, 0&, Process_ID _
               )

        If Handle_Process = 0 Then
            Error_Message = "Ошибка при попытке получить доступ к процессу: " & Err.LastDllError
            Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
            GoTo GT¦Next_PID
        End If

        With Power_Throttling
            .Version = 1
            .ControlMask = PROCESS_POWER_THROTTLING_EXECUTION_SPEED
            .StateMask = 0

            If Activate_EfficiencyMode Then
                Result_Priority = SetPriorityClass(Handle_Process, Low_Priority)
                .StateMask = PROCESS_POWER_THROTTLING_EXECUTION_SPEED
            Else
                Result_Priority = SetPriorityClass(Handle_Process, Normal_Priority)
            End If

            Result_EfficiencyMode = SetProcessInformation( _
                                    Handle_Process, Process_PowerThrottling, _
                                    Power_Throttling, LenB(Power_Throttling) _
                      )
        End With

        CloseHandle Handle_Process
        Threads_EfficiencyMode = (Result_Priority <> 0 And Result_EfficiencyMode <> 0)

        If Not Threads_EfficiencyMode Then
            Error_Message = "Ошибка при попытке установить режим эффективности: " & Err.LastDllError
            Show_ErrorMessage_Immediate Error_Message, "Ошибка доступа к процессу: " & Process_ID
            Final_Result = False
        End If

GT¦Next_PID:
    Next i

    Threads_EfficiencyMode = Final_Result
    '`````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------'
End Function
'====================================================================================================='


'======================================================================================================='
Public Function VBComponent_Build( _
                Optional ByVal VBComponent_Name As String = "modSDI_Template", _
                Optional ByVal VBComponent_ForcedReplacement As Boolean = True, _
                Optional ByVal dw_Reserved_1 As Long = 0&, _
                Optional ByVal dw_Reserved_2 As Long = 0&, _
                Optional ByVal dw_Reserved_3 As Long = 0& _
       ) As Boolean
'-------------------------------------------------------------------------------------------------------'
' // Documentation: In_Process
'```````````````````````````````````````````````````````````````````````````````````````````````````````'
' // < Зарезервированные параметры >

' dw_Reserved_1 - Use_DoEvents         | Добавить ли функкию DoEvents в генерируемый шаблон
' dw_Reserved_2 - Delay_InMilliseconds | Использовать ли задержку между циклами в генерируемом шаблоне
' dw_Reserved_3 - Type_VBTemplate      | Выбор готового шаблона для генерации
'-------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    MsgBox _
    "Данная функция (""VBComponent_Build"") находится в разработке!", vbInformation, "[DarkSec_Project]"
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'======================================================================'
Private Sub Init_VBD_Kit_Interface_SDI()
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


'====================================================================================================='
Private Function IsExcel_Responsive( _
                 ByRef Hwnd_Excel As Long, _
                 Optional ByVal Time_Out As Long = TIMEOUT_PROCESS _
        ) As Boolean
'-----------------------------------------------------------------------------------------------------'

    '````````````````````'
    Dim result As LongPtr
    '````````````````````'

    '```````````````````````````````````````````````````````````````'
    If Hwnd_Excel = 0 Then IsExcel_Responsive = False: Exit Function
    '```````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````'
    If CBool(SendMessageTimeout(Hwnd_Excel, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG, Time_Out, result)) Then
        IsExcel_Responsive = True
    Else
        IsExcel_Responsive = False
    End If
    '`````````````````````````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------'
End Function
'====================================================================================================='


'============================================================================='
Private Function CheckProcess_State( _
                 ByVal Process_ID As Long _
        ) As Process_State
'-----------------------------------------------------------------------------'

    '````````````````````````````````````````````'
    Dim Handle_Process As Long, Exit_Code As Long
    '````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Handle_Process = OpenProcess(NT_PROCESS_QUERY_INFORMATION, 0&, Process_ID)
    '`````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    If CBool(Handle_Process) Then
        If GetExitCodeProcess(Handle_Process, Exit_Code) Then
            If Exit_Code = STILL_ACTIVE Then
                CheckProcess_State = P_State_ActiveExecution
            Else
                CheckProcess_State = P_State_Terminate
            End If
        Else
            CheckProcess_State = P_State_ErrorReading
        End If
        Call CloseHandle(Handle_Process)
    Else
        CheckProcess_State = P_State_NonExistent
    End If
    '````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------'
End Function
'============================================================================='


'====================================================================='
Private Function KillProcess_PID( _
                 ByVal Process_ID As Long _
        ) As Boolean
'---------------------------------------------------------------------'

    '`````````````````````````````````````````'
    Dim Handle_Process As Long, result As Long
    '`````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    Handle_Process = OpenProcess(NT_PROCESS_TERMINATE, 0&, Process_ID)
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    If CBool(Handle_Process) Then
        result = TerminateProcess(Handle_Process, 0&)
        KillProcess_PID = CBool(result)
        Call CloseHandle(Handle_Process)
    Else
        KillProcess_PID = False
    End If
    '````````````````````````````````````````````````'

'---------------------------------------------------------------------'
End Function
'====================================================================='


'=================================================================================='
Private Function Get_ProcessPath_PID( _
                 ByVal Process_ID As Long _
        ) As String
'----------------------------------------------------------------------------------'

    '`````````````````````````````````````````````````````````````````'
    Dim Handle_Process As Long, result As Long, Buffer As String * 260
    '`````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    Handle_Process = _
    OpenProcess(NT_PROCESS_QUERY_INFORMATION Or NT_PROCESS_VM_READ, 0&, Process_ID)
    '``````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````'
    If Handle_Process <> 0 Then
        result = GetModuleFileNameEx(Handle_Process, 0, Buffer, 260)

        If result > 0 Then
            Get_ProcessPath_PID = Left$(Buffer, result)
        End If

        CloseHandle Handle_Process
    End If
    '```````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------'
End Function
'=================================================================================='


'=============================================================================='
Private Function Get_ProcessPriority_PID( _
                 ByVal Process_ID As Long _
        ) As Process_Priority
'------------------------------------------------------------------------------'

    '````````````````````````````'
    Dim Priority_Level As Long
    Dim Handle_Process As LongPtr
    '````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    Handle_Process = OpenProcess(NT_PROCESS_QUERY_INFORMATION, 0&, Process_ID)

    If CBool(Handle_Process) Then
        Priority_Level = GetPriorityClass(Handle_Process)
        CloseHandle Handle_Process
        Select Case Priority_Level
            Case 32:    Get_ProcessPriority_PID = NORMAL_PRIORITY_CLASS
            Case 64:    Get_ProcessPriority_PID = IDLE_PRIORITY_CLASS
            Case 128:   Get_ProcessPriority_PID = HIGH_PRIORITY_CLASS
            Case 256:   Get_ProcessPriority_PID = REALTIME_PRIORITY_CLASS
            Case 16384: Get_ProcessPriority_PID = BELOW_NORMAL_PRIORITY_CLASS
            Case 32768: Get_ProcessPriority_PID = ABOVE_NORMAL_PRIORITY_CLASS
            Case Else:  Get_ProcessPriority_PID = Priority_Level
        End Select
    Else
        Get_ProcessPriority_PID = Unknown_Priority
    End If
    '``````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------'
End Function
'=============================================================================='


'======================================================================================================'
Private Function Get_RunningProcesses_ByName( _
                 ByVal Process_Name As String _
        ) As Variant
'------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````'
    Dim Processes_PID()     As Long
    Dim Count_NeedProcesses As Long
    Dim Handle_Snapshot     As LongPtr
    Dim Process_ExeName     As String, Ret As Long
    Dim ProcessEntry32      As WinAPI_ProcessEntry32
    '```````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    Get_RunningProcesses_ByName = Array(): ProcessEntry32.dwSize = LenB(ProcessEntry32)
    '``````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````'
    Handle_Snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If Handle_Snapshot = 0 Then Exit Function
    '````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````'
    Ret = Process32First(Handle_Snapshot, ProcessEntry32): ReDim Processes_PID(1 To 512) As Long
    If Not LCase$(Right$(Process_Name, 4)) = ".exe" Then Process_Name = Process_Name & ".exe"
    '```````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````````'
    Do While Ret
        Process_ExeName = Left$(ProcessEntry32.szExeFile, InStr(ProcessEntry32.szExeFile, Chr$(0)) - 1)

        If StrComp(Process_ExeName, Process_Name, vbTextCompare) = 0 Then
            Count_NeedProcesses = Count_NeedProcesses + 1
            If UBound(Processes_PID, 1) < Count_NeedProcesses Then
                ReDim Preserve Processes_PID(1 To Count_NeedProcesses)
            End If
            Processes_PID(Count_NeedProcesses) = ProcessEntry32.th32ProcessID
        End If

        Ret = Process32Next(Handle_Snapshot, ProcessEntry32)
    Loop
    '``````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````'
    Call CloseHandle(Handle_Snapshot)
    '````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    If Not Count_NeedProcesses = 0 Then
        ReDim Preserve Processes_PID(1 To Count_NeedProcesses)
        Get_RunningProcesses_ByName = Processes_PID
    End If
    '`````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'=============================================================='
Private Function Get_RunningProcess_ByHwnd( _
                 ByRef Hwnd_XlMain As Long _
        ) As String
'--------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Window_Title As String, Buffer As String * 512
    Dim Process_Name As String, Length As Long, Inx_Pos As Long
    '``````````````````````````````````````````````````````````'

    '```````````````````````````````````````'
    Get_RunningProcess_ByHwnd = vbNullString
    '```````````````````````````````````````'

    '````````````````````````````````````````````````'
    Length = GetWindowText(Hwnd_XlMain, Buffer, 512&)
    Window_Title = Left(Buffer, Length)
    '````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    Inx_Pos = InStr(Window_Title, " - ")

    If Inx_Pos > 0 Then
        Process_Name = Left(Window_Title, Inx_Pos - 1)
    Else
        Process_Name = Window_Title
    End If
    '````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    If IsNumeric(Right(Process_Name, 1)) Then
        If InStr(1, Process_Name, ".") = 0 Then
            Get_RunningProcess_ByHwnd = Process_Name
        End If
    End If
    '````````````````````````````````````````````````'

'--------------------------------------------------------------'
End Function
'=============================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'========================================================================================================='
Private Function Controller_Thread( _
                 ByRef Thread_ExecuteCommand As Long, _
                 Optional ByRef Thread_Parameter As String = vbNullString, _
                 Optional ByRef Ignore_RunTime As Boolean = True _
        ) As String
'---------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````'
    Dim VBComponent_Name    As String, VBComponent_Data  As String
    Dim Vector_VBComponents As Variant, Inx_LB1 As Long, Inx_UB1 As Long
    '```````````````````````````````````````````````````````````````````````'
    Dim Obj_VBProject  As Object, Obj_VBComponent As Object, i As Long
    Dim Obj_CodeModule As Object, Line_Num As Long, Line_Text  As String
    '```````````````````````````````````````````````````````````````````````'
    Dim Flag_KernelUpdate    As Boolean, Thread_State       As Process_State
    Dim Flag_ProcedureExists As Boolean, Metrics_Module     As Variant
    '```````````````````````````````````````````````````````````````````````'
    Dim Module_Header_Delimiter As String, Module_Delimiter As String
    '```````````````````````````````````````````````````````````````````````'
    
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    On Error GoTo GT¦Error_Handler

    Module_Delimiter = Chr$(95) & Chr$(47) & "_" & Chr$(92) & Chr$(95)
    Module_Header_Delimiter = Chr$(95) & Chr$(92) & "_" & Chr$(47) & Chr$(95)

    Select Case Thread_ExecuteCommand
        Case CommandThread_Initialize:                GoSub GS¦CommandThread_Initialize
        Case CommandThread_Terminate:                 GoSub GS¦CommandThread_Terminate
        Case CommandThread_Run:                       GoSub GS¦CommandThread_Run
        Case CommandThread_Halt:                      GoSub GS¦CommandThread_Halt
        Case CommandThread_AbortExecution:            GoSub GS¦CommandThread_AbortExecution
        Case CommandThread_VBProject_AddComponent:    GoSub GS¦CommandThread_VBProject_AddComponent
        Case CommandThread_VBProject_UpdateComponent: GoSub GS¦CommandThread_VBProject_UpdateComponent
        Case CommandThread_VBProject_RemoveComponent: GoSub GS¦CommandThread_VBProject_RemoveComponent
        Case Else
            ThisWorkbook.CustomDocumentProperties("ExecuteCommand_Status") = StatusCommandThread_NotFound
            Controller_Thread = Err.Number & "_" & StatusCommandThread_NotFound
            On Error GoTo 0: Exit Function
    End Select
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````'
    ThisWorkbook.CustomDocumentProperties("ExecuteCommand_ErrorCode") = Err.Number
    ThisWorkbook.CustomDocumentProperties("ExecuteCommand_Status") = StatusCommandThread_Done

    Controller_Thread = Err.Number & "_" & StatusCommandThread_Done

    On Error GoTo 0: Exit Function
    '````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````'
GS¦CommandThread_Initialize:

    If Thread_Parameter = vbNullString Then Thread_Parameter = "Main"
    GoSub GS¦Check_StateThread: GoSub GS¦Check_ProcedureExists

    If Flag_ProcedureExists Then
        Call Thread_Initialize(Thread_Parameter)
    Else
        Err.Raise 1004
    End If

    Return
'`````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````'
GS¦CommandThread_Terminate:

    Application.DisplayAlerts = False: Application.Quit

    Return
'```````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````````````'
GS¦CommandThread_AbortExecution:

    ThisWorkbook. _
    CustomDocumentProperties("ExecuteCommand_Status") = StatusCommandThread_Done

    ThisWorkbook. _
    CustomDocumentProperties("ExecuteCommand_ErrorCode") = Err.Number

    ThisWorkbook. _
    CustomDocumentProperties("State_Process") = P_State_AbortExecution: End

    Return
'````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦CommandThread_Run:

    If Thread_Parameter = vbNullString Then Thread_Parameter = "Main"
    GoSub GS¦Check_StateThread: GoSub GS¦Check_ProcedureExists

    If Flag_ProcedureExists Then
        Deferred_ProcedureCall Thread_Parameter, 0.2
        ThisWorkbook.CustomDocumentProperties("State_Process") = P_State_ActiveExecution
    Else
        Err.Raise 1004
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````'
GS¦CommandThread_Halt:

    ThisWorkbook. _
    CustomDocumentProperties("State_Process") = P_State_Idle
    
    Glb_Thread_HaltExecution = True

    Return
'````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦CommandThread_VBProject_AddComponent:
GS¦CommandThread_VBProject_UpdateComponent:

    GoSub GS¦Check_StateThread

    If CBool(InStr(1, Thread_Parameter, Module_Header_Delimiter)) Then
        Vector_VBComponents = Split(Thread_Parameter, Module_Header_Delimiter)
    Else
        Vector_VBComponents = Array(Thread_Parameter)
    End If

    Inx_LB1 = LBound(Vector_VBComponents, 1)
    Inx_UB1 = UBound(Vector_VBComponents, 1)

    ReDim Vector_VBComponent_Name(Inx_LB1 To Inx_LB1)
    ReDim Vector_VBComponent_Data(Inx_LB1 To Inx_LB1)

    For i = Inx_LB1 To Inx_UB1
        ReDim Preserve Vector_VBComponent_Name(Inx_LB1 To i)
        ReDim Preserve Vector_VBComponent_Data(Inx_LB1 To i)

        If CBool(InStr(1, Vector_VBComponents(i), Module_Delimiter)) Then
            Metrics_Module = Split(Vector_VBComponents(i), Module_Delimiter)
            Vector_VBComponent_Name(i) = Metrics_Module(0)
            Vector_VBComponent_Data(i) = Metrics_Module(1)
            VBComponent_Name = Vector_VBComponent_Name(i)
        Else
            Vector_VBComponent_Name(i) = Vector_VBComponents(i)
            Vector_VBComponent_Data(i) = vbNullChar

            VBComponent_Name = Mid$(Vector_VBComponents(i), _
                                    InStrRev(Vector_VBComponents(i), "\") + 1)
            VBComponent_Name = Mid$(VBComponent_Name, 1, _
                                    InStrRev(VBComponent_Name, ".") - 1)
        End If

        If Thread_ExecuteCommand = CommandThread_VBProject_AddComponent Then
            If IsModule_Exists(VBComponent_Name) Then Err.Raise 9
        End If
    Next i

    Flag_KernelUpdate = False

    If Thread_ExecuteCommand = CommandThread_VBProject_UpdateComponent Then
        For i = Inx_LB1 To Inx_UB1
            VBComponent_Name = Vector_VBComponent_Name(i)
            If IsModule_Exists(VBComponent_Name) Then
                If VBComponent_Name = VBComponent_KernelName Then
                    Flag_KernelUpdate = True
                    With ActiveWorkbook.VBProject.VBComponents(VBComponent_Name)
                        .Name = VBComponent_KernelName & "_" & VBComponent_RuntimeName
                    End With
                Else
                    With ActiveWorkbook.VBProject
                        .VBComponents.Remove .VBComponents(VBComponent_Name)
                    End With
                    Select Case Thread_State
                        Case P_State_ActiveExecution, P_State_UserTriggeredExecution
                        Case Else
                            If IsModule_Exists(VBComponent_Name) Then Err.Raise 9
                    End Select
                End If
            End If
        Next i
    End If

    For i = Inx_LB1 To Inx_UB1
        If Vector_VBComponent_Data(i) = vbNullChar Then
            ThisWorkbook.VBProject.VBComponents.Import Vector_VBComponent_Name(i)
        Else
            Call Create_ModuleWithCode(CStr(Vector_VBComponent_Name(i)), _
                                       CStr(Vector_VBComponent_Data(i)))
        End If
    Next i

    If Flag_KernelUpdate Then
        Call Deferred_ProcedureCall("Update_KernelThread_Code")
    End If

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````'
GS¦CommandThread_VBProject_RemoveComponent:
    
    GoSub GS¦Check_StateThread
    Vector_VBComponents = Split(Thread_Parameter, "|")

    Inx_LB1 = LBound(Vector_VBComponents, 1)
    Inx_UB1 = UBound(Vector_VBComponents, 1)

    ReDim Vector_Inx(Inx_LB1 To Inx_UB1)

    For i = Inx_LB1 To Inx_UB1
        If Not Get_SemanticDataType(Vector_VBComponents(i)) = MSDI_VBComponent_SDI Then
            Err.Raise 9
        End If
    Next i

    For i = Inx_LB1 To Inx_UB1
        With ThisWorkbook.VBProject
            .VBComponents.Remove .VBComponents(Vector_VBComponents(i))
            Select Case Thread_State
                Case P_State_ActiveExecution, P_State_UserTriggeredExecution
                Case Else
                    If IsModule_Exists(CStr(Vector_VBComponents(i))) Then Err.Raise 1004
            End Select
        End With

        If Vector_VBComponents(i) = _
           ThisWorkbook.CustomDocumentProperties("Executable_Component") Then
           ThisWorkbook.CustomDocumentProperties("Executable_Component") = vbNullChar
        End If
    Next i

    Return
'```````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Check_StateThread:

    Thread_State = ThisWorkbook.CustomDocumentProperties("State_Process")

    If Not Ignore_RunTime Then
        Select Case Thread_State
            Case P_State_ActiveExecution, P_State_UserTriggeredExecution
                With ThisWorkbook
                    .CustomDocumentProperties("ExecuteCommand_ErrorCode") = Err.Number
                    .CustomDocumentProperties("ExecuteCommand_Status") = StatusCommandThread_NeedAbort
                End With

                Controller_Thread = Err.Number & "_" & StatusCommandThread_NeedAbort
                On Error GoTo 0: Exit Function
        End Select
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````'
GS¦Check_ProcedureExists:

    Flag_ProcedureExists = False: Set Obj_VBProject = Application.VBE.ActiveVBProject

    For Each Obj_VBComponent In Obj_VBProject.VBComponents
        Set Obj_CodeModule = Obj_VBComponent.CodeModule: Line_Num = 1
        Do While Line_Num <= Obj_CodeModule.CountOfLines
            Line_Text = Obj_CodeModule.Lines(Line_Num, 1)
            If Line_Text Like "*Function " & Thread_Parameter & "*" Or _
               Line_Text Like "*Sub " & Thread_Parameter & "*" Then
               Line_Text = Trim(Line_Text)
                If Not Line_Text Like "'*" And _
                   Not Line_Text Like "REM*" Then
                    If Line_Text Like "Public " & Thread_Parameter & "*" Or _
                       Line_Text Like "Sub " & Thread_Parameter & "*" Or _
                       Line_Text Like "Function " & Thread_Parameter & "*" Then
                       Flag_ProcedureExists = True: Exit Do
                    End If
                End If
            End If
            Line_Num = Line_Num + 1
        Loop
        If Flag_ProcedureExists Then Exit For
    Next Obj_VBComponent

    Set Obj_CodeModule = Nothing: Set Obj_VBComponent = Nothing
    Set Obj_VBProject = Nothing

    Return
'`````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````'
GT¦Error_Handler:

    ThisWorkbook.CustomDocumentProperties("ExecuteCommand_ErrorCode") = Err.Number
    ThisWorkbook.CustomDocumentProperties("ExecuteCommand_Status") = StatusCommandThread_Error

    Controller_Thread = Err.Number & "_" & StatusCommandThread_Error: On Error GoTo 0
'`````````````````````````````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================='


'======================================================================================='
Private Sub Thread_Initialize( _
            Optional Name_Procedure As String = "Main" _
        )
'---------------------------------------------------------------------------------------'
    If IsProcedure_Exists(Name_Procedure) Then
        Glb_Thread_HaltExecution = False: Glb_Thread_StateExecution = False
        Call Deferred_ProcedureCall(Name_Procedure)
        ThisWorkbook.CustomDocumentProperties("State_Process") = P_State_ActiveExecution
    Else
        ThisWorkbook.CustomDocumentProperties("State_Process") = P_State_Idle
    End If
'---------------------------------------------------------------------------------------'
End Sub
'======================================================================================='


'========================================================================='
Private Sub Thread_Completed( _
            Optional ByVal dw_Reserved As LongPtr = 0& _
        )
'-------------------------------------------------------------------------'
    With ThisWorkbook
        .CustomDocumentProperties("State_Process") = P_State_Completed
        If .CustomDocumentProperties("Thread_WaitingAfterExecution") Then
            Application.DisplayAlerts = False: Application.Quit
        End If
    End With
'-------------------------------------------------------------------------'
End Sub
'========================================================================='


'============================================================='
Private Sub Update_KernelThread_Code( _
            Optional ByRef Delay_InMilliseconds As Long = 1 _
        )
'-------------------------------------------------------------'

    '```````````````````````````````````````````````'
    On Error Resume Next: Sleep Delay_InMilliseconds
    '```````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    With ActiveWorkbook.VBProject
        .VBComponents.Remove .VBComponents( _
                              VBComponent_KernelName & "_" & _
                              VBComponent_RuntimeName _
                      )
    End With
    '`````````````````````````````````````````````````````````'

'-------------------------------------------------------------'
End Sub
'============================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'============================================================================================================'
Private Function Get_SemanticDataType( _
                 ByRef Data As Variant _
        ) As MSOffice_Type_ContentFormat_SDI
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
        Case vbBoolean:              Get_SemanticDataType = MSDI_Bool_SDI:        Exit Function
        Case vbEmpty:                Get_SemanticDataType = MSDI_Empty_SDI:       Exit Function
        Case vbNull:                 Get_SemanticDataType = MSDI_Null_SDI:        Exit Function
        Case vbError, vbObjectError: Get_SemanticDataType = MSDI_Undefined_SDI:   Exit Function
        Case vbUserDefinedType:      Get_SemanticDataType = MSDI_UserDefType_SDI: Exit Function
    End Select
    '``````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If (Data_VarType And vbArray) And (Not Data_TypeName = "Range") Then
        Get_SemanticDataType = MSDI_Array_SDI: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_TypeName
        Case "String"
            If Len(Data) = 0 Then Get_SemanticDataType = MSDI_NullString_SDI: Exit Function
            If IsNumeric(Data) Then Get_SemanticDataType = MSDI_String_SDI:   Exit Function
            Data_Type = Get_Directory_Type(CStr(Data))

            Select Case Data_Type
                Case DirType_File:     Get_SemanticDataType = MSDI_File_SDI:                  Exit Function
                Case DirType_Folder:   Get_SemanticDataType = MSDI_Folder_SDI:                Exit Function
                Case DirType_NotFound: Get_SemanticDataType = MSDI_NonExistent_Directory_SDI: Exit Function
                Case DirType_Invalid:
                    On Error Resume Next
                    VBProject_Name = Application.VBE.ActiveVBProject.Name

                    If Not Err.Number = 0 Then
                        Error_Message = "Невозможно проверить тип данных (Компонент это или VB-Проект)! " & _
                                        "Тип данных будет определен как строка (String)!"
                        Call Show_ErrorMessage_Immediate(Error_Message, "Проблема идентификации типов")
                        Get_SemanticDataType = MSDI_String_SDI: On Error GoTo 0: Exit Function
                    End If

                    If VBProject_Name = Data Then
                        Get_SemanticDataType = MSDI_VBProject_SDI
                        On Error GoTo 0: Exit Function
                    End If

                    Err.Clear
                    VBComponent_Name = Application.VBE.ActiveVBProject.VBComponents.Item(Data).Name
                    
                    If Len(VBComponent_Name) > 0 And Err.Number = 0 Then
                        Get_SemanticDataType = MSDI_VBComponent_SDI
                        On Error GoTo 0: Exit Function
                    End If

                    For Each Obj_VBProjects In Application.VBE.VBProjects
                        If Obj_VBProjects.Name = Data Then
                            Get_SemanticDataType = MSDI_VBProject_SDI:   On Error GoTo 0: Exit Function
                        End If

                        Err.Clear: Set Obj_VBComponent = Obj_VBProjects.VBComponents.Item(Data)
                        If (Err.Number = 0) And (Not Obj_VBComponent Is Nothing) Then
                            If Not Len(Obj_VBComponent.Name) = 0 Then
                                Set Obj_VBComponent = Nothing
                                Get_SemanticDataType = MSDI_VBComponent_SDI: On Error GoTo 0: Exit Function
                            End If
                        End If
                        Set Obj_VBComponent = Nothing
                    Next Obj_VBProjects

                    Get_SemanticDataType = MSDI_String_SDI: On Error GoTo 0: Exit Function
            End Select

        Case "Byte":       Get_SemanticDataType = MSDI_Int8_SDI:             Exit Function
        Case "Integer":    Get_SemanticDataType = MSDI_Int16_SDI:            Exit Function
        Case "Long":       Get_SemanticDataType = MSDI_Int32_SDI:            Exit Function
        Case "LongLong":   Get_SemanticDataType = MSDI_Int64_SDI:            Exit Function
        Case "Single":     Get_SemanticDataType = MSDI_FloatPoint32_SDI:     Exit Function
        Case "Currency":   Get_SemanticDataType = MSDI_FloatPoint64Cur_SDI:  Exit Function
        Case "Double":     Get_SemanticDataType = MSDI_FloatPoint64Dbl_SDI:  Exit Function
        Case "Decimal":    Get_SemanticDataType = MSDI_FloatPoint112_SDI:    Exit Function
        Case "Date":       Get_SemanticDataType = MSDI_Date_SDI:             Exit Function
        Case "Range":      Get_SemanticDataType = MSDI_Range_SDI:            Exit Function
        Case "Dictionary": Get_SemanticDataType = MSDI_Dictionary_SDI:       Exit Function
        Case "Collection": Get_SemanticDataType = MSDI_Collection_SDI:       Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````'
    If IsObject(Data) Then
        If Not Data Is Nothing Then
            Get_SemanticDataType = MSDI_Object_SDI:  Exit Function
        Else
            Get_SemanticDataType = MSDI_Nothing_SDI: Exit Function
        End If
    Else
        Get_SemanticDataType = MSDI_Undefined_SDI:   Exit Function
    End If
    '`````````````````````````````````````````````````````````````'

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


'=================================================================================='
Private Sub Deferred_ProcedureCall( _
            ByRef Name_Macro As String, _
            Optional ByVal Delay_InSeconds As Double = 0.1 _
        )
'----------------------------------------------------------------------------------'

    '````````````````````````````````````````````````'
    If Delay_InSeconds < 0 Then Delay_InSeconds = 0.1
    '````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    If Name_Macro Like "*.*" Then
        If Not Left$(Name_Macro, 1) = "!" Then Name_Macro = "!" & Name_Macro
    End If
    '```````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    ExecuteExcel4Macro "On.Time(Now()+" & _
    Replace(CStr(Delay_InSeconds / 100024), ",", ".") & ", """ & Name_Macro & """)"
    '``````````````````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------'
End Sub
'=================================================================================='


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
'------------------------------------------------------------------'

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

'------------------------------------------------------------------'
End Function
'=================================================================='


'========================================================================================='
Private Function Get_Function_InModule( _
                 ByRef Std_Module As Object, _
                 ByRef Code_Module As Object, _
                 ByRef Name_Function As String _
        ) As String
'-----------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim Vector_CodeFunction() As String, Module_TotalLines As Long
    Dim Inx As Long, Inx_Line As Long, Inx_Start As Long, Inx_End As Long
    Dim Module_CurrentProcedure As String, i As Long
    '````````````````````````````````````````````````````````````````````'
    Const Std_Module_ID As Long = 1
    '````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    If Not Std_Module.Type = Std_Module_ID Then
        Get_Function_InModule = vbNullString: Exit Function
    End If
    '``````````````````````````````````````````````````````'

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


'========================================================================================='
Private Function Get_Sub_InModule( _
                 ByRef Std_Module As Object, _
                 ByRef Code_Module As Object, _
                 ByRef Name_Function As String _
        ) As String
'-----------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    Dim Vector_CodeFunction() As String, Module_TotalLines As Long
    Dim Inx As Long, Inx_Line As Long, Inx_Start As Long, Inx_End As Long
    Dim Module_CurrentProcedure As String, i As Long
    '````````````````````````````````````````````````````````````````````'
    Const Std_Module_ID As Long = 1
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    If Not Std_Module.Type = Std_Module_ID Then
        Get_Sub_InModule = vbNullString: Exit Function
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
                Inx_End = Get_CodeRow_InModule(Code_Module, Inx_Start, "End Sub")
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

    '```````````````````````````````````````````````````'
    Get_Sub_InModule = Join(Vector_CodeFunction, vbCrLf)
    '```````````````````````````````````````````````````'

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


'============================================================================='
Private Function Get_Enums_InModule( _
                 ByRef Code_Module As Object, _
                 ByRef Declare_Enum As String _
        ) As String
'-----------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Inx As Long, Inx_Start As Long, Inx_End As Long
    '``````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    #If Has_PtrSafe Then
        Inx = Get_CodeRow_InModule(Code_Module, 1&, "Enum " & Declare_Enum)
        Inx_Start = Get_CodeRow_InModule(Code_Module, Inx, "End Enum")
    #Else
        Inx = Get_CodeRow_InModule(Code_Module, 1&, Declare_WinAPI)
        Inx = Get_CodeRow_InModule(Code_Module, Inx, Declare_WinAPI)
        Inx_Start = Get_CodeRow_InModule(Code_Module, Inx, "Declare", True)
    #End If

    If Declare_Enum Like "Function" Then
        Inx_End = Get_CodeRow_InModule(Code_Module, Inx, ") As")
    Else
        Inx_End = Inx_Start
    End If
    '`````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    Get_Enums_InModule = Get_Text_InModule(Code_Module, Inx - 3, Inx_End)
    Get_Enums_InModule = Remove_LeadingChars( _
                         Get_Enums_InModule, Chr$(10) & Chr$(13) & Chr$(32))
    '`````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------'
End Function
'============================================================================='


'========================================================================'
Private Function Get_Text_InModule( _
                 ByRef Code_Module As Object, _
                 ByRef Inx_StartLine As Long, _
                 ByRef Inx_EndLine As Long _
        ) As String
'------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Vector_CodeFunction() As String, Inx As Long, i As Long
    '``````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````'
    ReDim Vector_CodeFunction(1 To Inx_EndLine - Inx_StartLine)

    For i = Inx_StartLine To Inx_EndLine - 1
        Inx = Inx + 1: Vector_CodeFunction(Inx) = Code_Module.Lines(i, 1)
    Next i
    '````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    Get_Text_InModule = Join(Vector_CodeFunction, vbCrLf)
    '````````````````````````````````````````````````````'

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


'================================================================'
Private Function Get_ModuleCode( _
                 ByRef Module_Name As String _
       ) As String
'----------------------------------------------------------------'

    '``````````````````````````````````````````````````````'
    Dim Obj_VBProject  As Object, Obj_VBComponent As Object
    Dim Obj_CodeModule As Object, Total_Lines     As Long
    '``````````````````````````````````````````````````````'
    
    '````````````````````````````'
    Get_ModuleCode = vbNullString
    '````````````````````````````'
    
    '````````````````````````````````````````````````````````````'
    Set Obj_VBProject = ThisWorkbook.VBProject
    Set Obj_VBComponent = Obj_VBProject.VBComponents(Module_Name)
    Set Obj_CodeModule = Obj_VBComponent.CodeModule
    '````````````````````````````````````````````````````````````'
    
    '````````````````````````````````````````````````````````'
    Total_Lines = Obj_CodeModule.CountOfLines
    
    If Total_Lines > 0 Then
        Get_ModuleCode = Obj_CodeModule.Lines(1, Total_Lines)
    End If
    '````````````````````````````````````````````````````````'
    
'----------------------------------------------------------------'
End Function
'================================================================'


'================================================================================='
Private Sub Create_ModuleWithCode( _
            ByRef Module_Name As String, _
            ByRef Module_Code As String _
        )
'---------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````'
    Dim Obj_VBProject  As Object, Obj_VBComponent As Object
    Dim Component_Code As String
    '``````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    Set Obj_VBProject = ThisWorkbook.VBProject
    Set Obj_VBComponent = Obj_VBProject.VBComponents.Add(1)
    '``````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    With Obj_VBComponent
        .Name = Module_Name
        .CodeModule.AddFromString Module_Code
        With .CodeModule
            Component_Code = .Lines(1, .CountOfLines)
            If Right$(Component_Code, 2) = "()" Then .DeleteLines (.CountOfLines)
            If InStr(1, Component_Code, "Option Explicit") = 1 Then .DeleteLines 1
            If .Lines(1, 1) = vbNullString Then .DeleteLines 1
        End With
    End With
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````'
    Set Obj_VBProject = Nothing: Set Obj_VBComponent = Nothing
    '`````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------------'
End Sub
'================================================================================='


'========================================================================='
Private Function IsModule_Exists( _
                 ByRef Module_Name As String _
        ) As Boolean
'-------------------------------------------------------------------------'

    '````````````````````````````'
    Dim Obj_VBComponent As Object
    '````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    On Error Resume Next

    Set Obj_VBComponent = ThisWorkbook.VBProject.VBComponents(Module_Name)
    IsModule_Exists = (Err.Number = 0): Set Obj_VBComponent = Nothing

    On Error GoTo 0
    '`````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------'
End Function
'========================================================================='


'======================================================================================================'
Private Function IsProcedure_Exists( _
                 ByRef Procedure_Name As String _
        ) As Boolean
'------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````'
    Dim Obj_VBProject As Object, Obj_VBComponent  As Object
    Dim Full_Code    As Variant, Trimmed_Line     As String
    '``````````````````````````````````````````````````````'
    Dim Line_Count    As Long, i             As Long
    Dim Inx_1         As Long, Inx_2         As Long
    Dim Inx_3         As Long, Inx_4         As Long
    Dim Length_1      As Long, Length_2      As Long
    Dim Length_3      As Long, Length_4      As Long
    Dim SubString_1   As String, SubString_2 As String
    Dim SubString_3   As String, SubString_4 As String
    Dim SubString_5   As String, SubString_6 As String
    '``````````````````````````````````````````````````````'
    Static Obj_RegExp As Object
    '``````````````````````````````````````````````````````'

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

    '``````````````````````````````````````````````````````````````````````````````````````````````````'
    Set Obj_VBProject = ThisWorkbook.VBProject

    For Each Obj_VBComponent In Obj_VBProject.VBComponents
        With Obj_VBComponent.CodeModule
            Line_Count = .CountOfLines
            If Line_Count = 0 Then GoTo GS¦Next_VBComponent
            Full_Code = .Lines(1, .CountOfLines)
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

GS¦Next_VBComponent:
    Next Obj_VBComponent

    IsProcedure_Exists = False: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````````````````````'

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

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'=========================================================================='
Private Sub Sorting_VBComponents( _
            ByRef Source_Array() As Variant, _
            ByVal Low_Index As Long, _
            ByVal High_Index As Long _
        )
'--------------------------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim Pivot_Index As Long, Mid_Index  As Long
    Dim tmp_Val As String, i As Long, N As Long
    Dim Pivot_Value As String
    '``````````````````````````````````````````'

    '```````````````````````````````````````'
    If Low_Index >= High_Index Then Exit Sub
    '```````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````'
    Do While Low_Index < High_Index
        GoSub GS¦Partition

        If Pivot_Index - Low_Index < High_Index - Pivot_Index Then
            Sorting_VBComponents Source_Array, Low_Index, Pivot_Index - 1
            Low_Index = Pivot_Index + 1
        Else
            Sorting_VBComponents Source_Array, Pivot_Index + 1, High_Index
            High_Index = Pivot_Index - 1
        End If
    Loop

    Exit Sub
    '``````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````'
GS¦Partition:

    Mid_Index = (Low_Index + High_Index) \ 2

    If StrComp(Source_Array(Mid_Index), Source_Array(Low_Index)) < 0 Then
        tmp_Val = Source_Array(Mid_Index)
        Source_Array(Mid_Index) = Source_Array(Low_Index)
        Source_Array(Low_Index) = tmp_Val
    End If

    If StrComp(Source_Array(High_Index), Source_Array(Low_Index)) < 0 Then
        tmp_Val = Source_Array(High_Index)
        Source_Array(High_Index) = Source_Array(Low_Index)
        Source_Array(Low_Index) = tmp_Val
    End If

    If StrComp(Source_Array(High_Index), Source_Array(Mid_Index)) < 0 Then
        tmp_Val = Source_Array(High_Index)
        Source_Array(High_Index) = Source_Array(Mid_Index)
        Source_Array(Mid_Index) = tmp_Val
    End If

    tmp_Val = Source_Array(Mid_Index)
    Source_Array(Mid_Index) = Source_Array(High_Index - 1)
    Source_Array(High_Index - 1) = tmp_Val

    Pivot_Value = tmp_Val: i = Low_Index - 1: N = High_Index

    Do
        Do: i = i + 1: Loop While i < High_Index - 1 And _
                                  StrComp(Source_Array(i), Pivot_Value) < 0

        Do: N = N - 1: Loop While N > Low_Index And _
                                  StrComp(Source_Array(N), Pivot_Value) > 0

        If i < N Then
            tmp_Val = Source_Array(i)
            Source_Array(i) = Source_Array(N)
            Source_Array(N) = tmp_Val
        Else
            Exit Do
        End If
    Loop

    Pivot_Index = i: tmp_Val = Source_Array(i)
    Source_Array(i) = Source_Array(High_Index - 1)
    Source_Array(High_Index - 1) = tmp_Val

    Return
'``````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------'
End Sub
'=========================================================================='


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

    '``````````````````````````````'
    Call Init_VBD_Kit_Interface_SDI
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
