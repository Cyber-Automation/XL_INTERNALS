Attribute VB_Name = "VBD_Kit_UnitTest_Framework"
' | ========================================================================================================================= | '
' | __________  __________________________     ____________  _____________________  ___________________________________   __  | '
' | __  ____/ \/ /__  __ )__  ____/__  __ \    ___    |_  / / /__  __/_  __ \__   |/  /__    |__  __/___  _/_  __ \__  | / /  | '
' | _  /    __  /__  __  |_  __/  __  /_/ /    __  /| |  / / /__  /  _  / / /_  /|_/ /__  /| |_  /   __  / _  / / /_   |/ /   | '
' | / /___  _  / _  /_/ /_  /___  _  _, _/     _  ___ / /_/ / _  /   / /_/ /_  /  / / _  ___ |  /   __/ /  / /_/ /_  /|  /    | '
' | \____/  /_/  /_____/ /_____/  /_/ |_|______/_/  |_\____/  /_/    \____/ /_/  /_/  /_/  |_/_/    /___/  \____/ /_/ |_/     | '
' |                                     _/_____/                                                                              | '
' | ========================================================================================================================= | '

' +-[MODULE: VBD_Kit_UnitTest_Framework]--------------------------------------------------------------+
' |                                                                                                   |
' | [ENGINEER]: Zeus_0x01                                                                             |
' | [TELEGRAM]: @Zeus_0x01 (Public Name)                                                              |
' | [DESCRIPTION]: Реализация фреймворка для проведения тестирования различных процедур и компонентов |
' |                                                                                                   |
' +---------------------------------------------------------------------------------------------------+

' // <copyright file="VBD_Kit_UnitTest_Framework.bas" division="Cyber_Automation">
' // (C) Copyright 2025 Zeus_0x01 "{408DD863-D49F-48E3-BFAC-741C2E249C66}"
' // </copyright>

'-----------------------------------------------------------------------'
' // Implemented Functionality (Реализованный функционал):
'    Forum -
'    GitHub - https://github.com/Cyber-Automation/XL_INTERNALS/
'-----------------------------------------------------------------------'
' // Release_Version (Версия компонента) - [00.91]
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
Private Const GUID_VBComponent As String = "{408DD863-D49F-48E3-BFAC-741C2E249C66}"
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

'-------------------------------------------------'
Private Enum MSOffice_Type_ContentFormat_UnitTest
    MSDI_Undefined_UnitTest = &HFFFFFFFC
    MSDI_Null_UnitTest = &HFFFFFFFD
    MSDI_NullString_UnitTest = &HFFFFFFFE
    MSDI_Nothing_UnitTest = &HFFFFFFFF
    MSDI_Empty_UnitTest = &H0
    MSDI_Bool_UnitTest = &H1
    MSDI_Int8_UnitTest = &H2
    MSDI_Int16_UnitTest = &H4
    MSDI_Int32_UnitTest = &H8
    MSDI_Int64_UnitTest = &H10
    MSDI_FloatPoint32_UnitTest = &H20
    MSDI_FloatPoint64Cur_UnitTest = &H40
    MSDI_FloatPoint64Db_UnitTest = &H56
    MSDI_FloatPoint112_UnitTest = &H80
    MSDI_Date_UnitTest = &H100
    MSDI_String_UnitTest = &H200
    MSDI_Range_UnitTest = &H400
    MSDI_Array_UnitTest = &H800
    MSDI_Object_UnitTest = &H1000
    MSDI_Collection_UnitTest = &H2000
    MSDI_Dictionary_UnitTest = &H4000
    MSDI_File_UnitTest = &H8000
    MSDI_Folder_UnitTest = &H10000
    MSDI_VBComponent_UnitTest = &H20000
    MSDI_VBProject_UnitTest = &H40000
    MSDI_UserDefType_UnitTest = &H80000
    MSDI_NonExistent_Directory_UnitTest = &H100000
    
    ' {
        MSDI_Procedure_UnitTest = &H120000
        MSDI_AutoDetect_UnitTest = &HFFFFFFAD
    ' }
End Enum
'-------------------------------------------------'

'--------------------------------------'
Private Enum FileSystem_Directory_Type
    DirType_File = &H0
    DirType_Folder = &H1
    DirType_Invalid = &HFFFFFFFF
    DirType_NotFound = &HFFFFFFFE
End Enum
'--------------------------------------'

'---------------------------'
Private Enum OutPut_Data
    OutPut_Array = &H0
    OutPut_Range = &H1
    OutPut_String = &H2
    OutPut_Collection = &H3
    OutPut_Dictionary = &H4
End Enum
'---------------------------'

'---------------------------'
Private Enum State_UnitTest
    Type_Failed = &H0
    Type_Success = &H1
    Type_Ignored = &H2
    Type_Recompile = &H3
End Enum
'---------------------------'

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

'------------------------------'
Private Enum UI_TypeOperation
    Type_Addition
    Type_Increment
    Type_Assign
    Type_Clear
End Enum
'------------------------------'

'-------------------------------------'
Private Type UnitTest_Item
    Name_Test       As String
    Description     As String
    Status_Test     As State_UnitTest
    Execution_Time  As Double
    Result_Function As Variant
    Show_Result     As Boolean
End Type
'-------------------------------------'

'-------------------------------------'
Private Type UnitTest_Group
    Name_GroupTest   As String
    Count_Tests      As Long
    List_SubTests()  As UnitTest_Item
    Inx_Test         As Long
    Successful_Tests As Long
    Failed_Tests     As Long
    Ignored_Tests    As Long
    Success_Rate     As Double
End Type
'-------------------------------------'

'-------------------------------------'
Private Type UnitTest_Main
    Header_Table     As String
    Operating_System As String
    OS_BitDepth      As String
    HostVBA_BitDepth As String
    Processor_Model  As String
    Total_RAM        As String
    List_TestGroup() As UnitTest_Group
    Inx_Group        As Long
End Type
'-------------------------------------'

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

'------------------------------------------'
Private Type SYSTEM_INFO
    wProcessorArchitecture      As Integer
    wReserved                   As Integer
    dwPageSize                  As Long
    lpMinimumApplicationAddress As LongPtr
    lpMaximumApplicationAddress As LongPtr
    dwActiveProcessorMask       As LongPtr
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    wProcessorLevel             As Integer
    wProcessorRevision          As Integer
End Type
'------------------------------------------'

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

'-----------------------------------------'
Private Enum MSOffice_Type_Document_Group
    mso_Excel = &H2
    mso_PowerPoint = &H3
    mso_Word = &H4
    mso_Access = &H5
    mso_Outlook = &H6
End Enum
'-----------------------------------------'

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

'------------------------------------------------------------------------------------------'
#If Windows_NT And Has_PtrSafe Then  ' // Windows API (Kernel32.dll)

    Private Declare PtrSafe _
            Sub GetNativeSystemInfo Lib "kernel32.dll" ( _
                ByRef lpSystemInfo As SYSTEM_INFO _
            )
    
    Private Declare PtrSafe _
            Function GetFrequency Lib "kernel32.dll" Alias "QueryPerformanceFrequency" ( _
                     ByRef cyFrequency As Currency _
            ) As Long
    
    Private Declare PtrSafe _
            Function GetTickCount Lib "kernel32.dll" Alias "QueryPerformanceCounter" ( _
                     ByRef cyTickCount As Currency _
            ) As Long
    
    Private Declare PtrSafe _
            Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                ByRef Destination As Any, _
                ByRef Source As Any, _
                ByVal Length As LongPtr _
            )
                
    Private Declare PtrSafe _
            Function FileTimeToSystemTime Lib "kernel32.dll" ( _
                     ByRef lpFileTime As WinAPI_FileTime, _
                     ByRef lpSystemTime As WinAPI_SystemTime _
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

#Else

    ' // Old MS Office or MacOS

#End If
'------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
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
'-------------------------------------------------------------------------------------'

'------------------------------------------'
Private Header_Table      As String
Private UI_vbFramework    As Object
Private Dict_Log          As Object
Private Inx_TestPos       As Long
'------------------------------------------'
Private Structure_UT_Main As UnitTest_Main
'------------------------------------------'

'-------------------------------------------'
Private Glb_Thread_StateExecution As Boolean
'------------------------------------------------------------------'
Private Glb_MSOffice_Type_Building    As MSOffice_Type_Building
Private Glb_MSOffice_Type_Application As MSOffice_Type_Application
'------------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
Private Const REGISTRY_SECTION_MSO    As String = "SOFTWARE\Microsoft\Office\"

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


'======================================================================================================================='
Public Function Run_Framework() As Object
'-----------------------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````````````````'
    Dim Dict_UnitTests     As Object, Obj_ExlApp As Object
    Dim Dict_GroupTests    As Object, Obj_FSO    As Object
    Dim Dict_ResultTests   As Object, Obj_NewWB  As Object
    Dim Dict_RandomNames   As Object, Obj_WS     As Object
    '``````````````````````````````````````````````````````````````````````'
    Dim Source_Type        As MSOffice_Type_ContentFormat_UnitTest
    Dim Obj_OutPutSheet    As Object, Flag_Sheet  As Boolean
    Dim System_Information As String, Type_OutPut As OutPut_Data
    '``````````````````````````````````````````````````````````````````````'
    Dim Count_Rows As Long, Inx_Test    As Long
    Dim Inx_Row    As Long, End_Row     As Long
    Dim Inx_Group  As Long, Inx_RndName As Long
    '``````````````````````````````````````````````````````````````````````'
    Dim Matrix_UnitTest() As Variant, Matrix_Group()             As Variant
    Dim D_Key   As Variant, D_Item  As Variant, D_ItemResult     As Variant
    Dim tmp_Val As Variant, Inx_Pos As Long, i As Long, Rnd_Name As String
    '``````````````````````````````````````````````````````````````````````'
    Dim Array_Dimensions As Long, Start_Row     As Long
    Dim Array_LBound1    As Long, Array_UBound1 As Long
    Dim Array_LBound2    As Long, Array_UBound2 As Long
    Dim Matrix_Output()  As Variant, Col_Index  As Long
    Dim Cell_Value       As Variant, Row_Index  As Long
    '``````````````````````````````````````````````````````````````````````'
    Dim Range_Data() As Variant, Range_Rows As Long, Range_Cols As Long
    Dim Coll_Count   As Long, Coll_Index    As Long, Coll_Item  As Variant
    Dim Dict_Count  As Long, Dict_Key    As Variant, Dict_Index As Long
    '``````````````````````````````````````````````````````````````````````'
    Dim Matrix_Range()      As Variant
    Dim Matrix_Dictionary() As Variant
    Dim Matrix_Collection() As Variant
    Dim Vector_RndNames()   As Variant
    Dim Flag_Configuration  As Boolean
    '``````````````````````````````````````````````````````````````````````'
    Static Flag_Recompile   As Boolean
    '``````````````````````````````````````````````````````````````````````'
    Const xl_Maximized  As Long = -4137
    Const WS_Name       As String = "Main_UnitTest"
    Const UnitTest_Name As String = "VBD_Kit_UnitTest_Framework"
    '``````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    #If Not Windows_NT Then
        Call Show_InfoMessage_UnixNotSupported: Exit Function
    #End If
    '````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````'
    Call Init_VBD_Kit_UnitTest_Framework
    Call Init_AccessObjectModel(Glb_MSOffice_Type_Application)
    '````````````````````````````````````````````````````````````````````````````````'
    If Glb_MSOffice_Type_Building = Type_MSOffice_Build_NotSupport Then Exit Function
    '````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Flag_Configuration = False

    If UI_vbFramework Is Nothing Then
        Set UI_vbFramework = Create_UI_tmpForm: Flag_Configuration = True
        Update_UI_CommandLine UI_vbFramework.Name, "Получение конфигурации ПК ..."
    End If
    '`````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````'
    Call Init_UnitTest(UnitTest_Name)
    '````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````'
    With Structure_UT_Main
        System_Information = Empty
        System_Information = System_Information & "Информация о системе:" & vbNewLine
        System_Information = System_Information & "Операционная система: "
        System_Information = System_Information & .Operating_System & " " & .OS_BitDepth & vbNewLine
        System_Information = System_Information & "Разрядность приложения: " & .HostVBA_BitDepth & vbNewLine
        System_Information = System_Information & "Процессор: " & .Processor_Model & vbNewLine
        System_Information = System_Information & "ОЗУ (RAM): " & .Total_RAM

        If Flag_Configuration Then
            Update_UI_SystemConfiguration UI_vbFramework.Name, .Processor_Model, .Operating_System & " " & _
                                                               .OS_BitDepth, .HostVBA_BitDepth, .Total_RAM
            Update_UI_CommandLine UI_vbFramework.Name, "Конфигурация обновлена"
        End If
    End With
    '````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````'
    If Not Flag_Recompile Then
        Update_UI_CommandLine UI_vbFramework.Name, _
                 "Сканирование модуля и выполнение тестов (" & WS_Name & ") ..."
    End If

    If Run_UnitTests = Type_Recompile Then Flag_Recompile = True: Exit Function
    '``````````````````````````````````````````````````````````````````````````'

    '```````````````````````````'
    Call Get_Dictionaries( _
             Dict_UnitTests, _
             Dict_GroupTests, _
             Dict_ResultTests _
         )
    '```````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    If Dict_UnitTests.Count = 0 Then
        Update_UI_CommandLine UI_vbFramework.Name, "Не найдены тесты для запуска!"
        Update_UI_CommandLine UI_vbFramework.Name, "Проверьте сигнатуры функций для тестов!"
        Remove_UI_tmpForm UI_vbFramework: Set Run_Framework = Obj_NewWB
        GoSub GS¦Clear_StaticData: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````'
    GoSub GS¦Create_RandomNames: GoSub GS¦CreateMatrix_UnitTest: GoSub GS¦CreateMatrix_Group
    '```````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Set Obj_ExlApp = CreateObject("Excel.Application")
    Set Obj_NewWB = Obj_ExlApp.Workbooks.Add
    Set Obj_WS = Obj_NewWB.Worksheets(1)
    
    GoSub GS¦Create_BasicSheet
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````````````````````````````'
    Update_UI_CommandLine UI_vbFramework.Name, "Тестирование группы тестов (" & WS_Name & ") - завершено!"
    Obj_NewWB.Application.Visible = True: Remove_UI_tmpForm UI_vbFramework: Set Run_Framework = Obj_NewWB
    '`````````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````'
    GoSub GS¦Clear_StaticData: Exit Function
    '```````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_BasicSheet:

    With Obj_WS
        .Name = WS_Name
        .Cells(1, 1) = Structure_UT_Main.Header_Table
        .Range(.Cells(1, 1), .Cells(1, 7)).MergeCells = True

        With .Range("A1:G1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
                .PatternTintAndShade = 0
            End With

            With .Font
                .Bold = True: .Size = 14
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With

        .Cells(2, 1) = System_Information
        .Range(.Cells(2, 1), .Cells(6, 7)).MergeCells = True

        With .Range("A2:G6")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149998474074526
            End With

            With .Font
                .Bold = True
                .ThemeColor = xlThemeColorDark1
            End With
        End With

        .Cells(7, 1) = "№"
        .Cells(7, 2) = "Группа"
        .Cells(7, 3) = "Название теста"
        .Cells(7, 4) = "Результат"
        .Cells(7, 5) = "Описание"
        .Cells(7, 6) = "Время (сек.)"
        .Cells(7, 7) = "Результат процедуры"

        With .Range("A7:G7")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = 6299648
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
        End With

        Count_Rows = UBound(Matrix_UnitTest, 1)
        .Cells(8, 1).Resize(Count_Rows, UBound(Matrix_UnitTest, 2)).Value = Matrix_UnitTest

        With .Range("A7:G" & 7 + Count_Rows)
            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.499984740745262
                .Weight = xlThin
            End With
            .Font.Bold = True
        End With

        With .Range("A7:B" & 7 + Count_Rows)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("D7:D" & 7 + Count_Rows)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        For i = 8 To 7 + Count_Rows
            Select Case .Cells(i, "D").Value
            Case "Успешно": .Cells(i, "D").Interior.Color = 10867305
            Case "Ошибка":  .Cells(i, "D").Interior.Color = 3289855
            Case "Пропуск": .Cells(i, "D").Interior.Color = 65535
            End Select

            tmp_Val = .Cells(i, "G")
            If tmp_Val Like "EAX2D8//*" Then
                tmp_Val = Split(tmp_Val, "//")(1)

                If IsObject(Dict_ResultTests(tmp_Val)) Then
                    Set D_Item = Dict_ResultTests(tmp_Val)
                Else
                    D_Item = Dict_ResultTests(tmp_Val)
                End If

                GoSub GS¦Create_AdditionalSheet
                If Flag_Sheet Then
                    .Hyperlinks.Add Anchor:=.Cells(i, "G"), Address:="", _
                                    SubAddress:=Rnd_Name & "!A1", TextToDisplay:="LINK (SHEET)"
                Else
                    .Hyperlinks.Add Anchor:=.Cells(i, "G"), Address:=D_Item, TextToDisplay:="LINK (DIRECTORY)"
                End If
            End If
        Next i

        With .Range("F7:G" & 7 + Count_Rows)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Inx_Row = 8 + Count_Rows
        .Cells(Inx_Row + 1, 1) = "ИТОГОВАЯ СТАТИСТИКА ПО ГРУППАМ ТЕСТОВ"
        .Range(.Cells(Inx_Row + 1, 1), .Cells(Inx_Row + 1, 7)).MergeCells = True

        With .Range("A" & Inx_Row + 1 & ":G" & Inx_Row + 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
                .PatternTintAndShade = 0
            End With

            With .Font
                .Bold = True: .Size = 14
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With

        .Cells(Inx_Row + 2, 1) = "№"
        .Cells(Inx_Row + 2, 2) = "Группа тестов"
        .Range(.Cells(Inx_Row + 2, 2), .Cells(Inx_Row + 2, 3)).MergeCells = True

        .Cells(Inx_Row + 2, 4) = "Успешно"
        .Cells(Inx_Row + 2, 5) = "Ошибка"
        .Cells(Inx_Row + 2, 6) = "Пропущено"
        .Cells(Inx_Row + 2, 7) = "Процент успеха %"

        With .Range("A" & Inx_Row + 2 & ":G" & Inx_Row + 2)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = 6299648
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
        End With

        Count_Rows = UBound(Matrix_Group, 1)
        .Cells(Inx_Row + 3, 1).Resize(Count_Rows, UBound(Matrix_Group, 2)).Value = Matrix_Group
        End_Row = Inx_Row + 3 + Count_Rows

        With .Range("A" & Inx_Row + 3 & ":G" & End_Row)
            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.499984740745262
                .Weight = xlThin
            End With
            .Font.Bold = True
        End With

        With .Range("A" & Inx_Row + 3 & ":A" & End_Row)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("D" & Inx_Row + 3 & ":G" & End_Row)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("G" & Inx_Row + 3 & ":G" & End_Row)
            .NumberFormat = "0.00%"
        End With

        .Range("B" & Inx_Row + 3 & ":C" & End_Row).Merge True

        .Cells(End_Row, 1) = "ИТОГО:"
        .Cells(End_Row, 2) = "-"

        With .Cells(End_Row, 2)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range(.Cells(End_Row, 1), .Cells(End_Row, 7)).Borders
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlMedium
        End With

        .Cells(End_Row, 4).FormulaR1C1 = "=SUM(R[-" & Count_Rows & "]C:R[-1]C)"
        .Cells(End_Row, 5).FormulaR1C1 = "=SUM(R[-" & Count_Rows & "]C:R[-1]C)"
        .Cells(End_Row, 6).FormulaR1C1 = "=SUM(R[-" & Count_Rows & "]C:R[-1]C)"
        .Cells(End_Row, 7).FormulaR1C1 = "=SUM(R[-" & Count_Rows & "]C:R[-1]C)" & _
                                         "/COUNTA(R[-" & Count_Rows & "]C:R[-1]C)"

        .Calculate: .Columns("A:G").EntireColumn.AutoFit

        For i = Inx_Row + 3 To Inx_Row + 3 + Count_Rows
            If .Cells(i, "D").Value > 0 Then .Cells(i, "D").Interior.Color = 10867305
            If .Cells(i, "E").Value > 0 Then .Cells(i, "E").Interior.Color = 3289855
            If .Cells(i, "F").Value > 0 Then .Cells(i, "F").Interior.Color = 65535
        Next i
        
        .Range("H8").Activate
        With Obj_NewWB.Windows(1)
            .WindowState = xl_Maximized
            .FreezePanes = True
        End With
        .Range("A1").Activate
    End With

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦Create_AdditionalSheet:

    Source_Type = Get_SemanticDataType(D_Item): Flag_Sheet = False

    Select Case Source_Type
        Case MSDI_File_UnitTest:       Return
        Case MSDI_Folder_UnitTest:     Return
        Case MSDI_Range_UnitTest:      Type_OutPut = OutPut_Range
        Case MSDI_Array_UnitTest:      Type_OutPut = OutPut_Array
        Case MSDI_String_UnitTest:     Type_OutPut = OutPut_String
        Case MSDI_Collection_UnitTest: Type_OutPut = OutPut_Collection
        Case MSDI_Dictionary_UnitTest: Type_OutPut = OutPut_Dictionary
    End Select

    Inx_RndName = Inx_RndName + 1: Rnd_Name = "Result_" & Vector_RndNames(Inx_RndName)

    Set Obj_OutPutSheet = Obj_NewWB.Worksheets.Add(After:=Obj_NewWB.Worksheets(Obj_NewWB.Worksheets.Count))
    Obj_OutPutSheet.Name = Rnd_Name

    Select Case Type_OutPut
        Case OutPut_Array:      GoSub GS¦OutPut_ArrayToSheet
        Case OutPut_Range:      GoSub GS¦OutPut_RangeToSheet
        Case OutPut_String:     GoSub GS¦OutPut_StringToSheet
        Case OutPut_Collection: GoSub GS¦OutPut_CollectionToSheet
        Case OutPut_Dictionary: GoSub GS¦OutPut_DictionaryToSheet
    End Select

    Flag_Sheet = True: Obj_WS.Activate

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````'
GS¦CreateMatrix_UnitTest:

    ReDim Matrix_UnitTest(1 To Dict_UnitTests.Count, 1 To 7)

    Inx_Test = 0
    For Each D_Key In Dict_UnitTests.keys
        D_Item = Dict_UnitTests(D_Key): Inx_Test = Inx_Test + 1
        Matrix_UnitTest(Inx_Test, 1) = D_Item(1)
        Matrix_UnitTest(Inx_Test, 2) = D_Item(2)
        Matrix_UnitTest(Inx_Test, 3) = D_Item(3)
        Matrix_UnitTest(Inx_Test, 4) = D_Item(4)
        Matrix_UnitTest(Inx_Test, 5) = D_Item(5)
        Matrix_UnitTest(Inx_Test, 6) = D_Item(6)
        Matrix_UnitTest(Inx_Test, 7) = UCase$(D_Item(7))

        If IsObject(Dict_ResultTests(D_Key)) Then
            Set D_ItemResult = Dict_ResultTests(D_Key)
        Else
            D_ItemResult = Dict_ResultTests(D_Key)
        End If

        If Not IsEmpty(D_ItemResult) Then
            Matrix_UnitTest(Inx_Test, 7) = "EAX2D8//" & D_Key
        End If

    Next D_Key

    Return
'``````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````'
GS¦CreateMatrix_Group:

    ReDim Matrix_Group(1 To Dict_GroupTests.Count, 1 To 7)

    Inx_Group = 0
    For Each D_Key In Dict_GroupTests.keys
        D_Item = Dict_GroupTests(D_Key): Inx_Group = Inx_Group + 1
        Matrix_Group(Inx_Group, 1) = D_Item(1)
        Matrix_Group(Inx_Group, 2) = D_Item(2)
        Matrix_Group(Inx_Group, 3) = D_Item(3)
        Matrix_Group(Inx_Group, 4) = D_Item(4)
        Matrix_Group(Inx_Group, 5) = D_Item(5)
        Matrix_Group(Inx_Group, 6) = D_Item(6)
        Matrix_Group(Inx_Group, 7) = D_Item(7)
    Next D_Key

    Set Dict_GroupTests = Nothing: Set D_Item = Nothing

    Return
'``````````````````````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````````'
GS¦Create_RandomNames:

    Set Dict_RandomNames = CreateObject("Scripting.Dictionary")
    Set Obj_FSO = CreateObject("Scripting.FileSystemObject")

    For i = 1 To 16384
        Dict_RandomNames(Mid$(Obj_FSO.GetTempName, 4&, 5&)) = vbNullString
    Next i

    Vector_RndNames = Dict_RandomNames.keys
    Set Obj_FSO = Nothing

    Return
'``````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
GS¦OutPut_ArrayToSheet:

    With Obj_OutPutSheet
        .Hyperlinks.Add Anchor:=.Range("A1"), _
                        Address:="", SubAddress:=Obj_WS.Name & "!A1", _
                        TextToDisplay:="<<<"
        With .Range("A1")
            With .Font
                .Bold = True: .Size = 14
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("B1")
            .Value = "Результат (Массив):"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Font
                .Bold = True: .Size = 12
                .ThemeColor = xlThemeColorDark1
            End With

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
            End With
        End With

        Array_Dimensions = Get_ArrayDimension(D_Item)

        If Array_Dimensions = 0 Then
            .Range("D2").Value = "ОШИБКА: Массив не инициализирован"
            .Range("D2").Font.Color = RGB(255, 0, 0)
            .Columns("B:E").EntireColumn.AutoFit
        ElseIf Array_Dimensions = 1 Then
            Array_LBound1 = LBound(D_Item, 1)
            Array_UBound1 = UBound(D_Item, 1)
            
            If Array_UBound1 = -1 Then
                .Range("D2").Value = "ОШИБКА: Массив не инициализирован"
                .Range("D2").Font.Color = RGB(255, 0, 0)
                .Columns("B:E").EntireColumn.AutoFit
            Else
                ReDim Matrix_Output(1 To Array_UBound1 - Array_LBound1 + 1, 1 To 3)
    
                For Row_Index = Array_LBound1 To Array_UBound1
                    Cell_Value = Get_SafeCellValue(D_Item(Row_Index))
                    Inx_Pos = Row_Index - Array_LBound1 + 1
                    Matrix_Output(Inx_Pos, 1) = Inx_Pos
                    Matrix_Output(Inx_Pos, 2) = Cell_Value
                    Matrix_Output(Inx_Pos, 3) = "-"
                Next Row_Index
    
                .Range("D2").Value = "Индекс"
                .Range("E2").Value = "Значение"
                .Range("F2").Value = "-"
    
                With .Range("D2:F2")
                    .Interior.Color = 6299648
                    .Font.ThemeColor = xlThemeColorDark1
                    .Font.Bold = True
                End With
    
                With .Range("D3").Resize(Array_UBound1 - Array_LBound1 + 1, 3)
                    .Value = Matrix_Output
                End With
    
                .Columns("B:E").EntireColumn.AutoFit
    
                If .Columns("E:E").ColumnWidth > 100 Then
                    .Columns("E:E").ColumnWidth = 100
                End If
            End If

        ElseIf Array_Dimensions = 2 Then

            Array_LBound1 = LBound(D_Item, 1)
            Array_UBound1 = UBound(D_Item, 1)
            Array_LBound2 = LBound(D_Item, 2)
            Array_UBound2 = UBound(D_Item, 2)

            ReDim Matrix_Output(1 To Array_UBound1 - Array_LBound1 + 1, _
                                1 To Array_UBound2 - Array_LBound2 + 1)

            For Row_Index = Array_LBound1 To Array_UBound1
                For Col_Index = Array_LBound2 To Array_UBound2
                    Cell_Value = Get_SafeCellValue(D_Item(Row_Index, Col_Index))
                    Matrix_Output(Row_Index - Array_LBound1 + 1, _
                                  Col_Index - Array_LBound2 + 1) = Cell_Value
                Next Col_Index
            Next Row_Index

            Start_Row = 2
            For Col_Index = Array_LBound2 To Array_UBound2
                .Cells(Start_Row, Col_Index - Array_LBound2 + 4).Value = _
                                                                       "Столбец " & Col_Index
            Next Col_Index

            With .Range(.Cells(Start_Row, 4), _
                        .Cells(Start_Row, 4 + Array_UBound2 - Array_LBound2))
                .Interior.Color = 6299648
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
            End With

            .Range("D" & Start_Row + 1).Resize(UBound(Matrix_Output, 1), _
                                               UBound(Matrix_Output, 2)).Value = Matrix_Output

            .Columns("B:Z").EntireColumn.AutoFit
            
            For Col_Index = Array_LBound2 To Array_UBound2
                If .Columns(Col_Index + 3).ColumnWidth > 75 Then
                    .Columns(Col_Index + 3).ColumnWidth = 75
                End If
            Next Col_Index

        Else
            .Range("D2").Value = "ВНИМАНИЕ: Массив размерности " & Array_Dimensions & " - не поддерживается для вывода!"
            .Range("D3").Value = "Поддерживаются только одномерные и двумерные массивы!"
            .Range("D2:D3").Font.Color = RGB(255, 0, 0)
            .Columns("B:E").EntireColumn.AutoFit
        End If

        If Array_Dimensions > 0 And Array_Dimensions <= 2 Then
            With .Range("D2").CurrentRegion
                With .Borders
                    .LineStyle = xlContinuous
                    .ThemeColor = 1: .Weight = xlThin
                    .TintAndShade = -0.499984740745262
                End With
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If

        .Columns("C:C").ColumnWidth = 3.2

    End With

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````'
GS¦OutPut_RangeToSheet:

    With Obj_OutPutSheet
        .Hyperlinks.Add Anchor:=.Range("A1"), _
                        Address:="", SubAddress:=Obj_WS.Name & "!A1", _
                        TextToDisplay:="<<<"
        With .Range("A1")
            With .Font
                .Bold = True: .Size = 14
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("B1")
            .Value = "Результат (Диапазон):"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Font
                .Bold = True: .Size = 12
                .ThemeColor = xlThemeColorDark1
            End With

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
            End With
        End With

        If D_Item.Count = 1 Then
            Range_Data = Array(Array(Get_SafeCellValue(D_Item.Value)))
            Range_Rows = 1
            Range_Cols = 1
        Else
            Range_Data = D_Item.Value
            Range_Rows = D_Item.Rows.Count
            Range_Cols = D_Item.Columns.Count
        End If

        ReDim Matrix_Range(1 To Range_Rows, 1 To Range_Cols)

        For Row_Index = 1 To Range_Rows
            For Col_Index = 1 To Range_Cols
                If Range_Rows = 1 And Range_Cols = 1 Then
                    Cell_Value = Get_SafeCellValue(Range_Data(0)(0))
                Else
                    Cell_Value = Get_SafeCellValue(Range_Data(Row_Index, Col_Index))
                End If
                Matrix_Range(Row_Index, Col_Index) = Cell_Value
            Next Col_Index
        Next Row_Index

        .Range("D2").Value = "Адрес:"
        .Range("E2").Value = D_Item.Address(External:=True)
        .Range("D3").Value = "Размер:"
        .Range("E3").Value = Range_Rows & " x " & Range_Cols

        With .Range("D2:D3")
            .Interior.Color = 6299648
            .Font.ThemeColor = xlThemeColorDark1
        End With

        With .Range("D2:E3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1: .Weight = xlThin
                .TintAndShade = -0.499984740745262
            End With
            .Font.Bold = True
        End With

        .Range("G5").Resize(Range_Rows, Range_Cols).Value = Matrix_Range

        With .Range("G5").Resize(Range_Rows, Range_Cols)
            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1: .Weight = xlThin
                .TintAndShade = -0.499984740745262
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With .Range("G4")
            .Value = "Данные:"
            .Interior.Color = 6299648
            .Font.ThemeColor = xlThemeColorDark1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1: .Weight = xlThin
                .TintAndShade = -0.499984740745262
            End With
            .Font.Bold = True
        End With

        .Columns("B:E").EntireColumn.AutoFit
        .Columns("C:C").ColumnWidth = 3.2
        .Columns("F:F").ColumnWidth = 3.2

    End With

    Return
'```````````````````````````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GS¦OutPut_StringToSheet:

    With Obj_OutPutSheet
        .Hyperlinks.Add Anchor:=.Range("A1"), _
                        Address:="", SubAddress:=Obj_WS.Name & "!A1", _
                        TextToDisplay:="<<<"
        With .Range("A1")
            With .Font
                .Bold = True: .Size = 14
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("B1")
            .Value = "Результат (Строка):"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Font
                .Bold = True: .Size = 12
                .ThemeColor = xlThemeColorDark1
            End With

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
            End With
        End With

        .Range("D3") = Left$(CStr(D_Item), 32767&)
        .Range("D2").Value = "Значение"
        .Range("E2:E3").Value = "-"

        With .Range("D2:E2")
            .Interior.Color = 6299648
            .Font.ThemeColor = xlThemeColorDark1
        End With

        With .Range("D2:E3")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Borders
                .LineStyle = xlContinuous
                .ThemeColor = 1: .Weight = xlThin
                .TintAndShade = -0.499984740745262
            End With
        End With

        .Columns("B:D").EntireColumn.AutoFit

        If .Columns("D:D").ColumnWidth > 150 Then
            .Columns("D:D").ColumnWidth = 150
            .Range("D3").WrapText = True
        End If

        .Columns("C:C").ColumnWidth = 3.2

    End With

    Return
'````````````````````````````````````````````````````````````````````````'

'```````````````````````````````````````````````````````````````````````````````````````````````'
GS¦OutPut_CollectionToSheet:

    With Obj_OutPutSheet
        .Hyperlinks.Add Anchor:=.Range("A1"), _
                        Address:="", SubAddress:=Obj_WS.Name & "!A1", _
                        TextToDisplay:="<<<"
        With .Range("A1")
            With .Font
                .Bold = True: .Size = 14
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("B1")
            .Value = "Результат (Коллекция):"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Font
                .Bold = True: .Size = 12
                .ThemeColor = xlThemeColorDark1
            End With

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
            End With
        End With

        .Range("D2").Value = "Индекс"
        .Range("E2").Value = "Ключ"
        .Range("F2").Value = "Значение"
        .Range("G2").Value = "Тип данных"

        With .Range("D2:G2")
            .Interior.Color = 6299648
            .Font.ThemeColor = xlThemeColorDark1
            .Font.Bold = True
        End With

        Coll_Count = D_Item.Count

        If Coll_Count > 0 Then
            If Coll_Count > 1048572 Then
                .Range("D3").Value = "Невозможно вывести всю коллекцию (" & _
                                     Coll_Count & " строк) на лист!"
                .Range("D3:G3").Merge
                .Range("D3").HorizontalAlignment = xlCenter
            Else
                ReDim Matrix_Collection(1 To Coll_Count, 1 To 4)

                Coll_Index = 0
                For Each Coll_Item In D_Item
                    Coll_Index = Coll_Index + 1
                    Matrix_Collection(Coll_Index, 1) = Coll_Index
                    Matrix_Collection(Coll_Index, 2) = Get_KeyFromCollection(D_Item, Coll_Index)
                    Matrix_Collection(Coll_Index, 3) = Get_SafeCellValue(Coll_Item)
                    Matrix_Collection(Coll_Index, 4) = TypeName(Coll_Item)
                Next Coll_Item

                .Range("D3").Resize(Coll_Count, 4).Value = Matrix_Collection

                With .Range("D2:G" & 2 + Coll_Count)
                    With .Borders
                        .LineStyle = xlContinuous
                        .ThemeColor = 1
                        .Weight = xlThin
                        .TintAndShade = -0.499984740745262
                    End With
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If
        Else
            .Range("D3").Value = "Коллекция пуста"
            .Range("D3:G3").Merge
            .Range("D3").HorizontalAlignment = xlCenter
        End If

        .Columns("B:G").EntireColumn.AutoFit
        .Columns("C:C").ColumnWidth = 3.2
        
        If .Columns("E:E").ColumnWidth > 75 Then
            .Columns("E:E").ColumnWidth = 75
        End If
        
        If .Columns("F:F").ColumnWidth > 75 Then
            .Columns("F:F").ColumnWidth = 75
        End If

    End With

    Return
'```````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````'
GS¦OutPut_DictionaryToSheet:

    With Obj_OutPutSheet
        .Hyperlinks.Add Anchor:=.Range("A1"), _
                        Address:="", SubAddress:=Obj_WS.Name & "!A1", _
                        TextToDisplay:="<<<"
        With .Range("A1")
            With .Font
                .Bold = True: .Size = 14
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("B1")
            .Value = "Результат (Словарь):"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            With .Font
                .Bold = True: .Size = 12
                .ThemeColor = xlThemeColorDark1
            End With

            With .Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 4.99893185216834E-02
            End With
        End With

        .Range("D2").Value = "Ключ"
        .Range("E2").Value = "Значение"
        .Range("F2").Value = "Тип ключа"
        .Range("G2").Value = "Тип значения"

        With .Range("D2:G2")
            .Interior.Color = 6299648
            .Font.ThemeColor = xlThemeColorDark1
            .Font.Bold = True
        End With

        Dict_Count = D_Item.Count

        If Dict_Count > 0 Then
            If Dict_Count > 1048572 Then
                .Range("D3").Value = "Невозможно вывести весь словарь (" & _
                                     Dict_Count & " строк) на лист!"
                .Range("D3:G3").Merge
                .Range("D3").HorizontalAlignment = xlCenter
            Else
                ReDim Matrix_Dictionary(1 To Dict_Count, 1 To 4)

                Dict_Index = 0
                For Each Dict_Key In D_Item.keys
                    Dict_Index = Dict_Index + 1
                    Matrix_Dictionary(Dict_Index, 1) = Get_SafeCellValue(Dict_Key)
                    Matrix_Dictionary(Dict_Index, 2) = Get_SafeCellValue(D_Item(Dict_Key))
                    Matrix_Dictionary(Dict_Index, 3) = TypeName(Dict_Key)
                    Matrix_Dictionary(Dict_Index, 4) = TypeName(D_Item(Dict_Key))
                Next Dict_Key

                .Range("D3").Resize(Dict_Count, 4).Value = Matrix_Dictionary

                With .Range("D2:G" & 2 + Dict_Count)
                    With .Borders
                        .LineStyle = xlContinuous
                        .ThemeColor = 1
                        .Weight = xlThin
                        .TintAndShade = -0.499984740745262
                    End With
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If
        Else
            .Range("D3").Value = "Словарь пуст"
            .Range("D3:G3").Merge
            .Range("D3").HorizontalAlignment = xlCenter
        End If

        .Columns("B:G").EntireColumn.AutoFit
        .Columns("C:C").ColumnWidth = 3.2
        
        If .Columns("D:D").ColumnWidth > 75 Then
            .Columns("D:D").ColumnWidth = 75
        End If
        
        If .Columns("E:E").ColumnWidth > 75 Then
            .Columns("E:E").ColumnWidth = 75
        End If

    End With

    Return
'`````````````````````````````````````````````````````````````````````````````````````````'

'``````````````````````````'
GS¦Clear_StaticData:

    Set Dict_Log = Nothing
    Inx_TestPos = 0
    Flag_Recompile = False

    Return
'``````````````````````````'

'-----------------------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================================='


'======================================================='
Private Sub Internal_Call(): Call Run_Framework: End Sub
'======================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'


'==========================================================================='
Private Sub Init_VBD_Kit_UnitTest_Framework()
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


'======================================================================================'
Private Function Get_SafeCellValue( _
                 ByRef Value_Item As Variant _
        ) As Variant
'--------------------------------------------------------------------------------------'

    '````````````````````````````'
    On Error GoTo GT¦Handle_Error
    '````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    Select Case True

        Case IsError(Value_Item): Get_SafeCellValue = "Type_Name: Error": Exit Function
        Case IsNull(Value_Item):  Get_SafeCellValue = "Type_Name: Null":  Exit Function
        Case IsEmpty(Value_Item): Get_SafeCellValue = Value_Item:         Exit Function
    
        Case IsObject(Value_Item)
            If Value_Item Is Nothing Then
                Get_SafeCellValue = "Type_Name: Nothing"
            Else
                Get_SafeCellValue = "Type_Name: " & TypeName(Value_Item)
            End If
            Exit Function
    
        Case IsArray(Value_Item)
            Get_SafeCellValue = "Type_Name: Array(" & _
                                Get_ArrayDimension(Value_Item) & "D)": Exit Function

    End Select
    '``````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````'
    Select Case VarType(Value_Item)
        Case vbString
            If Len(Value_Item) = 0 Then
                Get_SafeCellValue = Value_Item
            Else
                Get_SafeCellValue = Left$(CStr(Value_Item), 32767&)
            End If
    
        Case vbBoolean, vbByte, vbInteger, vbLong, _
             vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate
            Get_SafeCellValue = Value_Item
    
        Case Else
            Get_SafeCellValue = "Type_Name: " & TypeName(Value_Item)

    End Select

    Exit Function
    '```````````````````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
GT¦Handle_Error:

    Get_SafeCellValue = "Type_Name: " & TypeName(Value_Item) & " (Error)"
    On Error GoTo 0
'````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------'
End Function
'======================================================================================'


'================================================================================================'
Private Function Run_UnitTests() As State_UnitTest
'------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Obj_VBComp As Object, Dict_UnitTests As Object
    Dim D_Key      As Variant, D_Item        As Variant
    Dim Inx_LB1    As Long, Inx_UB1 As Long, i  As Long
    Dim Current_Test   As Variant
    Dim Function_Name  As String
    Dim Test_Name      As String, tmp_Val As Variant
    Dim Flag_Result    As State_UnitTest
    Dim Execution_Time As Double
    Dim Inx_CountTests As Long
    '``````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Set Obj_VBComp = Find_ModuleByGUID(GUID_VBComponent)
    Set Dict_UnitTests = Get_AllProcedures_UnitTest(Obj_VBComp)

    If Dict_UnitTests.Count = 0 Then Exit Function
    '``````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````'
    If Dict_Log Is Nothing Then Set Dict_Log = CreateObject("Scripting.Dictionary")
    '``````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    For Each D_Key In Dict_UnitTests.keys
        D_Item = Dict_UnitTests(D_Key)
        Inx_CountTests = Inx_CountTests + UBound(D_Item, 1)
    Next D_Key
    '``````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    Call Update_UI_ControlProperty( _
         UI_vbFramework.Name, "L_Count_SummaryUT", _
         Inx_CountTests, Type_Assign _
                        )
    '``````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_UnitTests.keys
        D_Item = Dict_UnitTests(D_Key)
        Inx_LB1 = LBound(D_Item, 1)
        Inx_UB1 = UBound(D_Item, 1)

        With Structure_UT_Main
            .Inx_Group = .Inx_Group + 1

            If UBound(.List_TestGroup, 1) < .Inx_Group Then
                ReDim Preserve .List_TestGroup(1 To .Inx_Group + 32&)
            End If

            With .List_TestGroup(.Inx_Group)
                For i = Inx_LB1 To Inx_UB1
                    .Inx_Test = i

                    On Error Resume Next
                    tmp_Val = UBound(.List_SubTests, 1)
                    If Err.Number <> 0 Then ReDim .List_SubTests(1 To 1)
                    On Error GoTo 0

                    If UBound(.List_SubTests, 1) < .Inx_Test Then
                        ReDim Preserve .List_SubTests(1 To .Inx_Test + 32&)
                    End If

                    Current_Test = Split(D_Item(i), "//")
                    Function_Name = Current_Test(0)
                    Test_Name = Current_Test(1)
                    Flag_Result = Application.Run(Function_Name, i, Test_Name)

                    If Flag_Result = Type_Recompile Then
                        Run_UnitTests = Flag_Result: Exit Function
                    Else
                    End If

                    If Not Dict_Log.Exists(D_Key & "_" & .Inx_Test) Then

                        Inx_TestPos = Inx_TestPos + 1
                        Update_UI_ProgressBar UI_vbFramework.Name, Inx_TestPos, Inx_CountTests
                        Update_UI_ControlProperty UI_vbFramework.Name, _
                                                  "L_Count_FinishUT", 0&, Type_Increment

                        Select Case Flag_Result
                        Case Type_Success: Update_UI_ControlProperty UI_vbFramework.Name, _
                             "L_Count_SuccessUT", 0&, Type_Increment
                        Case Type_Failed:  Update_UI_ControlProperty UI_vbFramework.Name, _
                             "L_Count_FailedUT", 0&, Type_Increment

                        Case Type_Ignored: Update_UI_ControlProperty UI_vbFramework.Name, _
                             "L_Count_IgnoredUT", 0&, Type_Increment
                        End Select

                        Dict_Log(D_Key & "_" & .Inx_Test) = vbNullString: DoEvents
                        
                    End If

                Next i
                
                Erase Current_Test: ReDim Preserve .List_SubTests(1 To .Inx_Test)
                On Error Resume Next
                .Success_Rate = .Successful_Tests / .Count_Tests: DoEvents:
                If Err.Number <> 0 Then On Error GoTo 0: Exit Function
                On Error GoTo 0
                
                ' // ToDo
                
            End With
        End With: Erase D_Item
    Next D_Key
    '````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````'
    With Structure_UT_Main
        ReDim Preserve .List_TestGroup(1 To .Inx_Group)
        .Inx_Group = 0
    End With
    '``````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````'
    For Each D_Key In Dict_UnitTests.keys
        D_Item = Dict_UnitTests(D_Key)
        Inx_LB1 = LBound(D_Item, 1)
        Inx_UB1 = UBound(D_Item, 1)

        With Structure_UT_Main
            .Inx_Group = .Inx_Group + 1

            With .List_TestGroup(.Inx_Group)
                For i = Inx_LB1 To Inx_UB1
                    .Inx_Test = i: Current_Test = Split(D_Item(i), "//")
                    Function_Name = Current_Test(0): Test_Name = Current_Test(1)
                    Flag_Result = Application.Run(Function_Name, i, Test_Name, True)
                Next i
            End With
        End With: Erase D_Item
    Next D_Key
    '```````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````'
    Set Obj_VBComp = Nothing: Set Dict_UnitTests = Nothing
    '`````````````````````````````````````````````````````'

    '------------------------------------------------------------------------------------------------'
End Function

'================================================================================================'


'==========================================================================================================================================='
Private Sub Get_Dictionaries( _
            ByRef Dict_UnitTests As Object, _
            ByRef Dict_GroupTests As Object, _
            ByRef Dict_ResultTests As Object _
        )
'-------------------------------------------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````'
    Dim Source_Type As MSOffice_Type_ContentFormat_UnitTest
    Dim Inx_LB1     As Long, Inx_UB1 As Long
    Dim SubInx_UB1  As Long, Inx_Test  As Long
    Dim Inx_Group  As Long
    Dim i As Long, N As Long
    Dim Group_Name As String
    Dim Vector_UnitTest() As Variant
    Dim Vector_Group() As Variant, Show_Flag As Boolean
    '``````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Set Dict_UnitTests = CreateObject("Scripting.Dictionary")
    Set Dict_GroupTests = CreateObject("Scripting.Dictionary")
    Set Dict_ResultTests = CreateObject("Scripting.Dictionary")
    '``````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    With Structure_UT_Main
        Inx_LB1 = 1: Inx_UB1 = .Inx_Group: Inx_Test = 0: Inx_Group = 0

        For i = Inx_LB1 To Inx_UB1
            Inx_Group = Inx_Group + 1

            With .List_TestGroup(i)
                SubInx_UB1 = .Inx_Test: Group_Name = .Name_GroupTest

                For N = Inx_LB1 To SubInx_UB1
                    Inx_Test = Inx_Test + 1
                    With .List_SubTests(N)
                        ReDim Vector_UnitTest(1 To 7)
                        Vector_UnitTest(1) = Inx_Test
                        Vector_UnitTest(2) = Group_Name
                        Vector_UnitTest(3) = .Name_Test

                        Select Case .Status_Test
                            Case Type_Success: Vector_UnitTest(4) = "Успешно"
                            Case Type_Ignored: Vector_UnitTest(4) = "Пропуск"
                            Case Type_Failed:  Vector_UnitTest(4) = "Ошибка"
                        End Select

                        Vector_UnitTest(5) = .Description
                        Vector_UnitTest(6) = .Execution_Time

                        Source_Type = Get_SemanticDataType(.Result_Function)

                        Select Case Source_Type
                            Case MSDI_Undefined_UnitTest: Vector_UnitTest(7) = "Undefined":      Show_Flag = False
                            Case MSDI_Null_UnitTest:      Vector_UnitTest(7) = "Null":           Show_Flag = False
                            Case MSDI_Nothing_UnitTest:   Vector_UnitTest(7) = "Nothing":        Show_Flag = False
                            Case MSDI_Empty_UnitTest:     Vector_UnitTest(7) = "Empty":          Show_Flag = False
                            Case MSDI_Bool_UnitTest:      Vector_UnitTest(7) = .Result_Function: Show_Flag = False
    
                            Case MSDI_Int8_UnitTest To MSDI_Int64_UnitTest
                                Vector_UnitTest(7) = .Result_Function: Show_Flag = False
    
                            Case MSDI_FloatPoint32_UnitTest To MSDI_FloatPoint112_UnitTest
                                Vector_UnitTest(7) = .Result_Function: Show_Flag = False
    
                            Case MSDI_Date_UnitTest:                  Vector_UnitTest(7) = .Result_Function:         Show_Flag = False
                            Case MSDI_String_UnitTest:                Vector_UnitTest(7) = "String":                 Show_Flag = True
                            Case MSDI_Range_UnitTest:                 Vector_UnitTest(7) = "Range":                  Show_Flag = True
                            Case MSDI_Array_UnitTest:                 Vector_UnitTest(7) = "Array":                  Show_Flag = True
                            Case MSDI_Object_UnitTest:                Vector_UnitTest(7) = "Object":                 Show_Flag = False
                            Case MSDI_Collection_UnitTest:            Vector_UnitTest(7) = "Collection":             Show_Flag = True
                            Case MSDI_Dictionary_UnitTest:            Vector_UnitTest(7) = "Dictionary":             Show_Flag = True
                            Case MSDI_File_UnitTest:                  Vector_UnitTest(7) = "File":                   Show_Flag = True
                            Case MSDI_Folder_UnitTest:                Vector_UnitTest(7) = "Folder":                 Show_Flag = True
                            Case MSDI_VBComponent_UnitTest:           Vector_UnitTest(7) = "VB_Component":           Show_Flag = False
                            Case MSDI_VBProject_UnitTest:         Vector_UnitTest(7) = "VB_Project":             Show_Flag = False
                            Case MSDI_UserDefType_UnitTest:           Vector_UnitTest(7) = "User_Def_Type":          Show_Flag = False
                            Case MSDI_NonExistent_Directory_UnitTest: Vector_UnitTest(7) = "Non_Existent_Directory": Show_Flag = False
                        End Select

                        If Show_Flag Then
                            If .Show_Result Then
                                If IsObject(.Result_Function) Then
                                    Set Dict_ResultTests(Right("00000" & Inx_Test, 6) & Group_Name & Vector_UnitTest(3)) = .Result_Function
                                Else
                                    Dict_ResultTests(Right("00000" & Inx_Test, 6) & Group_Name & Vector_UnitTest(3)) = .Result_Function
                                End If
                            End If
                        End If
                    End With

                    Dict_UnitTests(Right("00000" & Inx_Test, 6) & Group_Name & Vector_UnitTest(3)) = Vector_UnitTest
                Next N

                ReDim Vector_Group(1 To 7)
                Vector_Group(1) = Inx_Group
                Vector_Group(2) = "Тестирование (" & Group_Name & ")"
                Vector_Group(4) = .Successful_Tests
                Vector_Group(5) = .Failed_Tests
                Vector_Group(6) = .Ignored_Tests
                Vector_Group(7) = .Success_Rate

                Dict_GroupTests(Right("00000" & Inx_Group, 6) & Group_Name) = Vector_Group
            End With

        Next i
    End With
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------------------------------------------'
End Sub
'==========================================================================================================================================='


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'


'==========================================================================='
Private Sub Fill_MainStructure( _
            ByRef Name_Group As String, _
            ByRef Description_UnitTest As String, _
            ByRef Name_UTest As String, _
            ByVal Execution_Time As Double, _
            ByRef Result_UnitTest As Variant, _
            ByVal Status_UnitTest As State_UnitTest, _
            ByVal Show_Result As Boolean _
        )
'---------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````'
    If Len(Description_UnitTest) > 100& Then
        Description_UnitTest = Left$(Description_UnitTest, 100&) & " ..."
    End If
    '````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    With Structure_UT_Main
        With .List_TestGroup(.Inx_Group)
            .Name_GroupTest = Name_Group: .Count_Tests = .Count_Tests + 1
            With .List_SubTests(.Inx_Test)
                .Description = Description_UnitTest
                .Name_Test = Name_UTest
                .Execution_Time = Execution_Time
                If IsObject(Result_UnitTest) Then
                    Set .Result_Function = Result_UnitTest
                Else
                    .Result_Function = Result_UnitTest
                End If
                .Status_Test = Status_UnitTest
                .Show_Result = Show_Result
            End With

            Select Case Status_UnitTest
                Case Type_Success: .Successful_Tests = .Successful_Tests + 1
                Case Type_Failed:  .Failed_Tests = .Failed_Tests + 1
                Case Type_Ignored: .Ignored_Tests = .Ignored_Tests + 1
            End Select
        End With
    End With
    '```````````````````````````````````````````````````````````````````````'

'---------------------------------------------------------------------------'
End Sub
'==========================================================================='


'======================================================================================================='
Private Function Get_AllProcedures_UnitTest( _
                 ByRef Obj_VBComponent As Object _
        ) As Object
'-------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````'
    Dim Obj_CodeModule As Object
    Dim Dict_Result    As Object
    Dim Name_Function  As String
    Dim Name_GroupUT   As String
    Dim Name_UnitTest  As String
    Dim Length_InitSign  As Long
    Dim Code_Module      As Variant
    Dim Line_Text        As String
    Dim Inx_InitSign     As Long
    Dim Inx_FinalSign    As Long
    Dim Inx_Pos As Long, D_Item, i As Long
    Dim Vector_Items()     As Variant
    Dim Function_Signature As String
    '``````````````````````````````````````````````````'
    Const Initial_Signature As String = "Run_UnitTest_"
    '``````````````````````````````````````````````````'

    '````````````````````````````````````````````````````'
    Set Obj_CodeModule = Obj_VBComponent.CodeModule
    Set Dict_Result = CreateObject("Scripting.Dictionary")
    '````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````'
    Length_InitSign = Len(Initial_Signature)
    Function_Signature = "Private Function " & Initial_Signature & "*(*" & _
                                              "ByVal Inx_CurrentTest As Long,*" & _
                                              "Optional Name_UTest As String = *" & _
                                              "Optional Reset_StaticVariable As Boolean = False*" & _
                                              "Optional Name_Group As String = *" & _
                                              "Optional Description_UnitTest As String = *" & _
                                  ") As State_UnitTest"
    '````````````````````````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````'
    With Obj_CodeModule
        Code_Module = .Lines(1, .CountOfLines)
    End With
    '`````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````'
    Code_Module = Replace(Code_Module, " _" & vbNewLine, vbNullString)
    Code_Module = Split(Code_Module, vbNewLine)
    '````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    For i = 1 To UBound(Code_Module, 1)
        Line_Text = Trim$(Code_Module(i))
        If Line_Text Like Function_Signature Then
            Inx_InitSign = InStr(1, Line_Text, Initial_Signature)
            Inx_FinalSign = InStr(Inx_InitSign, Line_Text, "(") - 1

            If Inx_InitSign <> 0 Or Inx_FinalSign <> 0 Then
                On Error Resume Next
                Name_GroupUT = Mid$( _
                               Line_Text, _
                               Inx_InitSign + Length_InitSign, _
                               Inx_FinalSign - (Inx_FinalSign - Inx_InitSign) _
                               )

                If Err.Number = 0 Then
                    Err.Clear: Inx_Pos = InStr(Name_GroupUT, "_")
                    If Inx_Pos <> 0 Then
                        Name_Function = Mid$(Line_Text, Inx_InitSign, Inx_FinalSign - Inx_InitSign + 1)
                        Name_GroupUT = Mid$(Name_GroupUT, 1, Inx_Pos - 1)
                        Name_UnitTest = Mid$(Name_Function, InStr(1, Name_Function, Name_GroupUT) + _
                                                           Len(Name_GroupUT) + 1)
                        If Err.Number = 0 Then
                            ReDim Vector_Items(1 To 1)
                            Vector_Items(1) = Name_Function & "//" & Name_UnitTest
                            Dict_Result.Add Name_GroupUT, Vector_Items
                            If Err.Number <> 0 Then
                                D_Item = Dict_Result(Name_GroupUT)
                                ReDim Preserve D_Item(1 To UBound(D_Item, 1) + 1)
                                D_Item(UBound(D_Item, 1)) = Name_Function & "//" & Name_UnitTest
                                Dict_Result(Name_GroupUT) = D_Item
                            End If
                        End If
                    End If
                End If

                On Error GoTo 0
            End If

        End If

    Next i
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````'
    Set Get_AllProcedures_UnitTest = Dict_Result
    '```````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


'====================================================================================='
Private Function Get_KeyFromCollection( _
                 ByRef Obj_Collection As Variant, _
                 ByVal Item_Index As Long, _
                 Optional View_CacheMemory As Boolean = False _
        ) As String
'-------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````'
    Dim Ptr_Collection As LongPtr, Inx_Count  As Long
    Dim Ptr_Null As LongPtr, Key As String, i As Long
    '````````````````````````````````````````````````'
    Static Cache_Memory As Object
    '````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````'
    #If x64_Soft Then
        Const Ptr_Size       As Long = 8&, Var_Size As Long = 24&
        Const Coll_PtrOffset As Long = 40&
    #Else
        Const Ptr_Size       As Long = 4&, Var_Size As Long = 16&
        Const Coll_PtrOffset As Long = 24&
    #End If

    Ptr_Null = 0&
    '````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````'
    If Obj_Collection Is Nothing Then Exit Function
    '``````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    If Cache_Memory Is Nothing Then
        Set Cache_Memory = CreateObject("Scripting.Dictionary")
    End If
    '``````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````'
    If View_CacheMemory Then
        GoSub GS¦Add_CollectionInCacheMemory
        If Not Get_KeyFromCollection = vbNullString Then Exit Function
        Get_KeyFromCollection = Cache_Memory(CStr(Ptr_Collection))(Item_Index)
        Exit Function
    End If
    '`````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````'
    If Item_Index < 1& Then Exit Function
    If Item_Index > Obj_Collection.Count Then Exit Function
    '``````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    Ptr_Collection = ObjPtr(Obj_Collection)

    For i = 1 To Item_Index
        CopyMemory Ptr_Collection, ByVal Ptr_Collection + Coll_PtrOffset, Ptr_Size
    Next
    '`````````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````````````'
    CopyMemory ByVal VarPtr(Key), ByVal Ptr_Collection + Var_Size, Ptr_Size
    Get_KeyFromCollection = Key:  CopyMemory ByVal VarPtr(Key), Ptr_Null, Ptr_Size

    Exit Function
    '`````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````'
GS¦Add_CollectionInCacheMemory:

    Ptr_Collection = ObjPtr(Obj_Collection)

    If Not Cache_Memory.Exists(CStr(Ptr_Collection)) Then
        Inx_Count = Obj_Collection.Count: ReDim Vector_Cache(1 To Inx_Count)

        For i = 1 To Inx_Count
            CopyMemory Ptr_Collection, ByVal Ptr_Collection + Coll_PtrOffset, Ptr_Size
            CopyMemory ByVal VarPtr(Key), ByVal Ptr_Collection + Var_Size, Ptr_Size
            Vector_Cache(i) = Key:  CopyMemory ByVal VarPtr(Key), Ptr_Null, Ptr_Size
        Next i

        Ptr_Collection = ObjPtr(Obj_Collection)
        Cache_Memory.Add CStr(Ptr_Collection), Vector_Cache
        Get_KeyFromCollection = Vector_Cache(Item_Index)
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------'
End Function
'====================================================================================='


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'


'==================================================================================='
Private Sub Init_UnitTest( _
            ByRef Name_UnitTest As String _
        )
'-----------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````'
    With Structure_UT_Main
        #If x64_Soft Then
            .HostVBA_BitDepth = "{x64}": .Inx_Group = 0
            If Len(.Header_Table) = 0 Then
                ReDim .List_TestGroup(1 To 64)
            End If
        #Else
            .HostVBA_BitDepth = "{x32}": .Inx_Group = 0
            If Len(.Header_Table) = 0 Then
                ReDim .List_TestGroup(1 To 32)
            End If
        #End If

        .Header_Table = Name_UnitTest & " - Тестирование - " & Now

        If Len(.Total_RAM) = 0 Then .Total_RAM = Get_DataRAM()
        If Len(.OS_BitDepth) = 0 Then .OS_BitDepth = Get_BitDepthOS()
        If Len(.Operating_System) = 0 Then .Operating_System = Get_TypeOS()
        If Len(.Processor_Model) = 0 Then .Processor_Model = Get_FullNameProcessor()
    End With
    '```````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------'
End Sub
'==================================================================================='


'==================================================='
Private Function Micro_Timer() As Double
'---------------------------------------------------'

    '```````````````````````````'
    Dim Ticks_Count  As Currency
    Static Frequency As Currency
    '```````````````````````````'

    '```````````````````````````````````````````'
    Micro_Timer = 0: GetTickCount Ticks_Count
    If Frequency = 0 Then GetFrequency Frequency
    '```````````````````````````````````````````'

    '````````````````````````````````````````'
    If Frequency <> 0 Then
        Micro_Timer = Ticks_Count / Frequency
    End If
    '````````````````````````````````````````'

'---------------------------------------------------'
End Function
'==================================================='


'================================================================================================================'
Private Function Get_SemanticDataType( _
                 ByRef Data As Variant _
        ) As MSOffice_Type_ContentFormat_UnitTest
'----------------------------------------------------------------------------------------------------------------'

    '````````````````````````````````````````````````````````````````````````````````'
    Dim Data_VarType  As Long, Data_TypeName    As String, Error_Message    As String
    Dim Data_Type     As FileSystem_Directory_Type, VBProjects              As Object
    Dim tmp_Component As Object, VBProject_Name As String, VBComponent_Name As String
    '````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````'
    Data_VarType = VarType(Data): Data_TypeName = TypeName(Data)
    '``````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_VarType
        Case vbBoolean:              Get_SemanticDataType = MSDI_Bool_UnitTest:        Exit Function
        Case vbEmpty:                Get_SemanticDataType = MSDI_Empty_UnitTest:       Exit Function
        Case vbNull:                 Get_SemanticDataType = MSDI_Null_UnitTest:        Exit Function
        Case vbError, vbObjectError: Get_SemanticDataType = MSDI_Undefined_UnitTest:   Exit Function
        Case vbUserDefinedType:      Get_SemanticDataType = MSDI_UserDefType_UnitTest: Exit Function
    End Select
    '```````````````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````'
    If (Data_VarType And vbArray) And (Not Data_TypeName = "Range") Then
        Get_SemanticDataType = MSDI_Array_UnitTest: Exit Function
    End If
    '```````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'
    Select Case Data_TypeName
        Case "String"
            If Len(Data) = 0 Then Get_SemanticDataType = MSDI_NullString_UnitTest: Exit Function
            If IsNumeric(Data) Then Get_SemanticDataType = MSDI_String_UnitTest:   Exit Function
            Data_Type = Get_Directory_Type(CStr(Data))
    
            Select Case Data_Type
                Case DirType_File:     Get_SemanticDataType = MSDI_File_UnitTest:                  Exit Function
                Case DirType_Folder:   Get_SemanticDataType = MSDI_Folder_UnitTest:                Exit Function
                Case DirType_NotFound: Get_SemanticDataType = MSDI_NonExistent_Directory_UnitTest: Exit Function
                Case DirType_Invalid:
                    On Error Resume Next
                    VBProject_Name = Application.VBE.ActiveVBProject.Name
        
                    If Not Err.Number = 0 Then
                        Error_Message = "Невозможно проверить тип данных (Компонент это или VB-Проект)! " & _
                                        "Тип данных будет определен как строка (String)!"
                        Call Show_ErrorMessage_Immediate(Error_Message, Err, "Проблема идентификации типов")
                        Get_SemanticDataType = MSDI_String_UnitTest: On Error GoTo 0: Exit Function
                    End If
        
                    If VBProject_Name = Data Then
                        Get_SemanticDataType = MSDI_VBProject_UnitTest
                        On Error GoTo 0: Exit Function
                    End If
        
                    VBComponent_Name = Application.VBE.ActiveVBProject.VBComponents.Item(Data).Name
                    If Len(VBComponent_Name) > 0 Then
                        Get_SemanticDataType = MSDI_VBComponent_UnitTest
                        On Error GoTo 0: Exit Function
                    End If
        
                    For Each VBProjects In Application.VBE.VBProjects
                        If VBProjects.Name = Data Then
                            Get_SemanticDataType = MSDI_VBProject_UnitTest:   On Error GoTo 0: Exit Function
                        End If
        
                        Err.Clear: Set tmp_Component = VBProjects.VBComponents.Item(Data)
                        If (Err.Number = 0) And (Not tmp_Component Is Nothing) Then
                            If Not Len(tmp_Component.Name) = 0 Then
                                Set tmp_Component = Nothing
                                Get_SemanticDataType = MSDI_VBComponent_UnitTest: On Error GoTo 0: Exit Function
                            End If
                        End If
                        Set tmp_Component = Nothing
                    Next VBProjects
    
                    Get_SemanticDataType = MSDI_String_UnitTest: On Error GoTo 0: Exit Function
            End Select

        Case "Byte":       Get_SemanticDataType = MSDI_Int8_UnitTest:             Exit Function
        Case "Integer":    Get_SemanticDataType = MSDI_Int16_UnitTest:            Exit Function
        Case "Long":       Get_SemanticDataType = MSDI_Int32_UnitTest:            Exit Function
        Case "LongLong":   Get_SemanticDataType = MSDI_Int64_UnitTest:            Exit Function
        Case "Single":     Get_SemanticDataType = MSDI_FloatPoint32_UnitTest:     Exit Function
        Case "Currency":   Get_SemanticDataType = MSDI_FloatPoint64Cur_UnitTest:  Exit Function
        Case "Double":     Get_SemanticDataType = MSDI_FloatPoint64Db_UnitTest:  Exit Function
        Case "Decimal":    Get_SemanticDataType = MSDI_FloatPoint112_UnitTest:    Exit Function
        Case "Date":       Get_SemanticDataType = MSDI_Date_UnitTest:             Exit Function
        Case "Range":      Get_SemanticDataType = MSDI_Range_UnitTest:            Exit Function
        Case "Dictionary": Get_SemanticDataType = MSDI_Dictionary_UnitTest:       Exit Function
        Case "Collection": Get_SemanticDataType = MSDI_Collection_UnitTest:       Exit Function
    End Select
    '````````````````````````````````````````````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    If IsObject(Data) Then
        If Not Data Is Nothing Then
            Get_SemanticDataType = MSDI_Object_UnitTest:  Exit Function
        Else
            Get_SemanticDataType = MSDI_Nothing_UnitTest: Exit Function
        End If
    Else
        Get_SemanticDataType = MSDI_Undefined_UnitTest:   Exit Function
    End If
    '``````````````````````````````````````````````````````````````````'

'----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================'


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


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'


'======================================='
Private Function Get_TypeOS() As String
'---------------------------------------'

    '````````````````````````````'
    #If Windows_NT Then
        Get_TypeOS = "Windows_NT"
    #ElseIf Mac Then
        Get_TypeOS = "MacOS"
    #Else
        Get_TypeOS = "Unknown_OS"
    #End If
    '````````````````````````````'

'---------------------------------------'
End Function
'======================================='


'================================================================'
Private Function Get_BitDepthOS() As String
'----------------------------------------------------------------'

    '`````````````````````````````````````'
    Dim Sys_Info As SYSTEM_INFO
    Const PROCESSOR_ARCHITECTURE_AMD64 = 9
    '`````````````````````````````````````'

    '```````````````````````````'
    GetNativeSystemInfo Sys_Info
    '```````````````````````````'

    '````````````````````````````````````````````````````````````'
    If Sys_Info. _
       wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
        Get_BitDepthOS = "{x64}"
    Else
        Get_BitDepthOS = "{x32}"
    End If
    '````````````````````````````````````````````````````````````'

'----------------------------------------------------------------'
End Function
'================================================================'


'==============================================================================='
Private Function Get_DataRAM() As String
'-------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_WMIService As Object
    Dim Coll_RAM       As Object
    Dim Obj_RAM        As Object
    Dim Total_RAM      As Double
    '```````````````````````````'

    '```````````````````````````````````````````````````````````````````````````'
    Set Obj_WMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set Coll_RAM = Obj_WMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    '```````````````````````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````````'
    For Each Obj_RAM In Coll_RAM
        Total_RAM = Total_RAM + (Obj_RAM.Capacity / 1024& / 1024& / 1024&)
    Next Obj_RAM
    '`````````````````````````````````````````````````````````````````````'

    '``````````````````````````````'
    Get_DataRAM = Total_RAM & " GB"
    '``````````````````````````````'

'-------------------------------------------------------------------------------'
End Function
'==============================================================================='


'==========================================================================='
Private Function Get_FullNameProcessor() As String
'---------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_WMIService As Object
    Dim Coll_CPU       As Object
    Dim Obj_CPU        As Object
    Dim CPU_Name       As String
    '```````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    Set Obj_WMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set Coll_CPU = Obj_WMIService.ExecQuery("SELECT * FROM Win32_Processor")
    '```````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    For Each Obj_CPU In Coll_CPU
        CPU_Name = Obj_CPU.Name: Exit For
    Next Obj_CPU
    '````````````````````````````````````'

    '```````````````````````````````'
    Get_FullNameProcessor = CPU_Name
    '```````````````````````````````'

'---------------------------------------------------------------------------'
End Function
'==========================================================================='


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'


'================================================================================================================='
Private Function Create_UI_tmpForm() As Object
'-----------------------------------------------------------------------------------------------------------------'

    '```````````````````````````````````````````````````````````````````````````````````````'
    Dim UI_UserForm As Object, UI_Control As Object, UI_Frame As Object, Inx_Replace As Long
    Dim UserForm_Property As UI_UserForm_Property, Control_Property As UI_Control_Property
    '```````````````````````````````````````````````````````````````````````````````````````'
    Dim Control_Type      As UI_Controls_Type
    Dim Control_Name      As String, Control_Caption    As String
    Dim Control_FontSize  As Double, Control_FontBold   As Boolean
    Dim Control_Left      As Double, Control_Top        As Double
    Dim Control_Width     As Double, Control_Height     As Double
    Dim Control_BackColor As Variant, Control_ForeColor As Variant
    '```````````````````````````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    UserForm_Property.UI_Caption = "UnitTest_vbFramework (Cyber_Automation)"
    UserForm_Property.UI_Width = 470.25
    UserForm_Property.UI_Height = 245.25
    UserForm_Property.UI_ShowModal = False
    UserForm_Property.UI_ShowBorders = True
    UserForm_Property.UI_BackColor = vbBlack
    UserForm_Property.UI_BorderColor = vbBlack

    Set UI_UserForm = UI_Add_NewForm(UserForm_Property)
    '```````````````````````````````````````````````````````````````````````'

    '``````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Command"
    Control_Caption = Empty
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 180
    Control_Width = 450
    Control_Height = 12
    Control_BackColor = vbBlack
    Control_ForeColor = &HFFFF00
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '``````````````````````````````````````````'

    '``````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_ProgressBar"
    Control_Caption = String$(35, ChrW$(11036))
    Control_FontSize = 14
    Control_Left = 6
    Control_Top = 192
    Control_Width = 450
    Control_Height = 24
    Control_BackColor = vbBlack
    Control_ForeColor = &HFFFF00
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '``````````````````````````````````````````'

    '``````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_InfoSystem"
    Control_Caption = "Информация о системе:"
    Control_FontSize = 9
    Control_Left = 240
    Control_Top = 10
    Control_Width = 222
    Control_Height = 15
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '``````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_InfoSystem_Processor"
    Control_Caption = "Процессор: " & vbNewLine & "-"
    Control_FontSize = 9
    Control_Left = 240
    Control_Top = 30
    Control_Width = 222
    Control_Height = 30
    Control_BackColor = vbBlack
    Control_ForeColor = vbGreen
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '`````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_InfoSystem_OS"
    Control_Caption = "Операционная система: " & vbNewLine & "-"
    Control_FontSize = 9
    Control_Left = 240
    Control_Top = 65
    Control_Width = 222
    Control_Height = 30
    Control_BackColor = vbBlack
    Control_ForeColor = vbGreen
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '```````````````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_InfoSystem_Soft"
    Control_Caption = "Разрядность приложения: " & vbNewLine & "-"
    Control_FontSize = 9
    Control_Left = 240
    Control_Top = 100
    Control_Width = 222
    Control_Height = 30
    Control_BackColor = vbBlack
    Control_ForeColor = vbGreen
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '``````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_InfoSystem_RAM"
    Control_Caption = "ОЗУ (RAM): " & vbNewLine & "-"
    Control_FontSize = 9
    Control_Left = 240
    Control_Top = 135
    Control_Width = 222
    Control_Height = 30
    Control_BackColor = vbBlack
    Control_ForeColor = vbGreen
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '````````````````````````````````````````````````'

    '``````````````````````````````````````````'
    Control_Type = Control_Frame
    Control_Name = "Frm_Metadata_UnitTests"
    Control_Caption = "Метаданные тестов"
    Control_FontSize = 10
    Control_Left = 6
    Control_Top = 6
    Control_Width = 220
    Control_Height = 162
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_Control
    '``````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_SummaryUT"
    Control_Caption = "Общее количество тестов: "
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 12
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Count_SummaryUT"
    Control_Caption = "0"
    Control_FontSize = 8
    Control_Left = 183
    Control_Top = 12
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '``````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_FinishUT"
    Control_Caption = "Количество завершенных тестов: "
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 36
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '``````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Count_FinishUT"
    Control_Caption = "0"
    Control_FontSize = 8
    Control_Left = 183
    Control_Top = 36
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Frame
    Control_Name = "Frm_ResultUT"
    Control_Caption = Empty
    Control_FontSize = 10
    Control_Left = 6
    Control_Top = 60
    Control_Width = 204
    Control_Height = 84
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_SuccessUT"
    Control_Caption = "Количество успешных тестов:"
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 12
    Control_Width = 156
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Count_SuccessUT"
    Control_Caption = "0"
    Control_FontSize = 8
    Control_Left = 175.5
    Control_Top = 12
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_FailedUT"
    Control_Caption = "Количество провальных тестов: "
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 36
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Count_FailedUT"
    Control_Caption = "0"
    Control_FontSize = 8
    Control_Left = 175.5
    Control_Top = 36
    Control_Width = 168
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbRed
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_IgnoredUT"
    Control_Caption = "Количество пропущенных тестов: "
    Control_FontSize = 8
    Control_Left = 6
    Control_Top = 60
    Control_Width = 156
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbWhite
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '`````````````````````````````````````````````````'
    Control_Type = Control_Label
    Control_Name = "L_Count_IgnoredUT"
    Control_Caption = "0"
    Control_FontSize = 8
    Control_Left = 175.5
    Control_Top = 60
    Control_Width = 24
    Control_Height = 18
    Control_BackColor = vbBlack
    Control_ForeColor = vbYellow
    Control_FontBold = True

    GoSub GS¦Fill_Control: GoSub GS¦Add_ControlToFrame
    '`````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````'
    UserForm_Property = Clear_UserForm_Property(): UserForms.Add(UI_UserForm.Name).Show
    Set Create_UI_tmpForm = UI_UserForm: Exit Function
    '``````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````'
GS¦Fill_Control:

    With Control_Property
        .UI_Control_Type = Control_Type
        .UI_Control_Name = Control_Name
        .UI_Control_Caption = Control_Caption
        .UI_Control_FontSize = Control_FontSize
        .UI_Control_Left = Control_Left
        .UI_Control_Top = Control_Top
        .UI_Control_Width = Control_Width
        .UI_Control_Height = Control_Height
        .UI_Control_BackColor = Control_BackColor
        .UI_Control_ForeColor = Control_ForeColor
        .UI_Control_FontBold = Control_FontBold
    End With

    Return
'`````````````````````````````````````````````````'

'``````````````````````````````````````````````````````````````````````'
GS¦Add_Control:

    If Control_Property.UI_Control_Type = Control_Frame Then
        Set UI_Frame = UI_Add_NewControl(UI_UserForm, Control_Property)
    Else
        Call UI_Add_NewControl(UI_UserForm, Control_Property)
    End If

    Control_Property = Clear_Control_Property()

    Return
'``````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````'
GS¦Add_ControlToFrame:

    If Control_Property.UI_Control_Type = Control_Frame Then
        Set UI_Frame = UI_Add_NewControlToFrame(UI_Frame, Control_Property)
    Else
        Call UI_Add_NewControlToFrame(UI_Frame, Control_Property)
    End If
    Control_Property = Clear_Control_Property()

    Return
'`````````````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------------------------------------'
End Function
'================================================================================================================='


'============================================================'
Private Function Remove_UI_tmpForm( _
            ByRef UI_UserForm As Variant _
        ) As Boolean
'------------------------------------------------------------'

    '``````````````````````````````````````````'
    Dim Obj_tmpForm As Object, Source_UserForm
    '``````````````````````````````````````````'

    '```````````````````````````````````````````````````'
    For Each Source_UserForm In VBA.UserForms
        If Source_UserForm.Name = UI_UserForm.Name Then
            Source_UserForm.Hide
            Set Source_UserForm = Nothing
            GoTo GT¦Remove_Form
        End If
    Next Source_UserForm
    '```````````````````````````````````````````````````'

'````````````````````````````````````````````````````````````'
GT¦Remove_Form:

    With Application.VBE.ActiveVBProject
        While Obj_tmpForm Is Nothing
            Set Obj_tmpForm = .VBComponents(UI_UserForm.Name)
                .VBComponents.Remove Obj_tmpForm: DoEvents
        Wend
    End With

    Set Obj_tmpForm = Nothing: Set UI_UserForm = Nothing
'````````````````````````````````````````````````````````````'

'------------------------------------------------------------'
End Function
'============================================================'


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


'======================================================================================================================='
Private Function UI_Add_NewControlToFrame( _
                 ByRef UI_Frame As Object, _
                 ByRef Control_Property As UI_Control_Property _
        ) As Object
'-----------------------------------------------------------------------------------------------------------------------'

    '```````````````````````````'
    Dim Obj_UI_Control As Object
    '```````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'
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
        Case Control_Frame:         Set Obj_UI_Control = UI_Frame.Controls.Add("Forms.Frame.1", "btnInsideFrame")
        Case Else:                  Set UI_Add_NewControlToFrame = Nothing: Exit Function
    End Select
    '```````````````````````````````````````````````````````````````````````````````````````````````````````````````````'

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
    
'-----------------------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================================='


'==================================================================================='
Private Function Update_UI_SystemConfiguration( _
                 ByRef UI_tmpForm_Name As String, _
                 ByRef Processor_Model As String, _
                 ByRef Operating_System As String, _
                 ByRef HostVBA_BitDepth As String, _
                 ByRef Total_RAM As String _
        ) As Boolean
'-----------------------------------------------------------------------------------'

    '``````````````````````````````'
    Dim UI_UserForm       As Object
    Dim Len_ProcessorModel  As Long
    Dim Len_OperatingSystem As Long
    Dim Len_HostVBABitDepth As Long
    Dim Len_TotalRAM        As Long
    Dim Max_Len As Long, i  As Long
    Dim T       As Double
    '``````````````````````````````'

    '````````````````````````````````````````````````'
    Len_ProcessorModel = Len(Processor_Model) + 13&
    Len_OperatingSystem = Len(Operating_System) + 26&
    Len_HostVBABitDepth = Len(HostVBA_BitDepth) + 24&
    Len_TotalRAM = Len(Total_RAM) + 13&
    '````````````````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Max_Len = Len_ProcessorModel
    If Max_Len < Len_TotalRAM Then Max_Len = Len_TotalRAM
    If Max_Len < Len_OperatingSystem Then Max_Len = Len_OperatingSystem
    If Max_Len < Len_HostVBABitDepth Then Max_Len = Len_HostVBABitDepth
    '``````````````````````````````````````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            For i = 1 To Max_Len
                If i <= Len(Processor_Model) Then
                    UI_UserForm.L_InfoSystem_Processor _
                    .Caption = "Процессор: " & vbNewLine & _
                               Left$(Processor_Model, i)
                End If

                If i <= Len(Operating_System) Then
                    UI_UserForm.L_InfoSystem_OS _
                    .Caption = "Операционная система: " & vbNewLine & _
                               Left$(Operating_System, i)
                End If

                If i <= Len(HostVBA_BitDepth) Then
                    UI_UserForm.L_InfoSystem_Soft _
                    .Caption = "Разрядность приложения: " & vbNewLine & _
                               Left$(HostVBA_BitDepth, i)
                End If

                If i <= Len(Total_RAM) Then
                    UI_UserForm.L_InfoSystem_RAM _
                    .Caption = "ОЗУ (RAM): " & vbNewLine & _
                               Left$(Total_RAM, i)
                End If

                T = Timer: While Timer < T + 0.03
                    DoEvents
                Wend
            Next i
            Update_UI_SystemConfiguration = True: Exit Function
        End If
    Next UI_UserForm

    On Error GoTo 0: Update_UI_SystemConfiguration = False
    '```````````````````````````````````````````````````````````````````````'

'-----------------------------------------------------------------------------------'
End Function
'==================================================================================='


'============================================================'
Private Function Update_UI_CommandLine( _
                 ByRef UI_tmpForm_Name As String, _
                 ByRef Text As String _
        ) As Boolean
'------------------------------------------------------------'

    '`````````````````````````'
    Dim UI_UserForm  As Object
    Dim i As Long, T As Double
    '`````````````````````````'

    '````````````````````````````````````````````````````````'
    On Error Resume Next
    
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            For i = 1 To Len(Text)
                UI_UserForm.L_Command.Caption = Left(Text, i)
                T = Timer
                While Timer < T + 0.02: DoEvents: Wend
            Next i
            T = Timer
            While Timer < T + 1: DoEvents: Wend
            Update_UI_CommandLine = True: Exit Function
        End If
    Next UI_UserForm

    On Error GoTo 0: Update_UI_CommandLine = False
    '````````````````````````````````````````````````````````'

'------------------------------------------------------------'
End Function
'============================================================'


'=================================================================================================='
Private Function Update_UI_ProgressBar( _
                 ByRef UI_tmpForm_Name As String, _
                 ByRef Index_Test As Long, _
                 ByRef Count_Test As Long _
        ) As Boolean
'--------------------------------------------------------------------------------------------------'

    '````````````````````````'
    Dim UI_UserForm As Object
    Dim Count_Symbols As Long
    Dim Percent_Line As Double
    Dim Count_Chars As Long
    '````````````````````````'

    '``````````````````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            Count_Symbols = Len(UI_UserForm.L_ProgressBar.Caption)
            Percent_Line = Index_Test / Count_Test: Count_Chars = Int(Count_Symbols * Percent_Line)
            UI_UserForm.L_ProgressBar.Caption = String$(Count_Chars, ChrW$(11035)) & _
                                                String$(Count_Symbols - Count_Chars, ChrW$(11036))
            Update_UI_ProgressBar = True: Exit Function
        End If
    Next UI_UserForm

    On Error GoTo 0: Update_UI_ProgressBar = False
    '``````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------'
End Function
'=================================================================================================='


'======================================================================================================='
Private Function Update_UI_ControlProperty( _
                 ByRef UI_tmpForm_Name As String, _
                 ByRef Control_Name As String, _
                 ByRef Data_Caption As Long, _
                 ByRef Type_Operation As UI_TypeOperation _
        ) As Boolean
'-------------------------------------------------------------------------------------------------------'

    '````````````````````````'
    Dim UI_UserForm As Object
    Dim Obj_Control As Object
    '````````````````````````'

    '````````````````````````````````'
    Update_UI_ControlProperty = False
    '````````````````````````````````'

    '```````````````````````````````````````````````````````````````````````````````````````````````````'
    On Error Resume Next
    
    For Each UI_UserForm In VBA.UserForms
        If UI_UserForm.Name = UI_tmpForm_Name Then
            On Error Resume Next
            Set Obj_Control = UI_UserForm.Controls(Control_Name)
            On Error GoTo 0

            If Not Obj_Control Is Nothing Then
                On Error Resume Next
                Select Case Type_Operation
                    Case Type_Addition:  Obj_Control.Caption = CLng(Obj_Control.Caption) + Data_Caption
                    Case Type_Increment: Obj_Control.Caption = CLng(Obj_Control.Caption) + 1
                    Case Type_Assign:    Obj_Control.Caption = Data_Caption
                    Case Type_Clear:     Obj_Control.Caption = Empty
                End Select
                If Err.Number = 0 Then Update_UI_ControlProperty = True
                On Error GoTo 0: Exit Function
            Else
                On Error GoTo 0: Exit Function
            End If
        End If
    Next UI_UserForm
    
    On Error GoTo 0
    '```````````````````````````````````````````````````````````````````````````````````````````````````'

'-------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================='


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
    
    '````````````````````````````````````'
    Call Init_VBD_Kit_UnitTest_Framework
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

' // <<< UT_Template >>>

''====================================================================================================='
'Private Function Run_UnitTest_{GroupName}_{TestName}( _
'                 ByVal Inx_CurrentTest As Long, _
'                 Optional Name_UTest As String = "-", _
'                 Optional Reset_StaticVariable As Boolean = False, _
'                 Optional Name_Group As String = {"NameGroup"}, _
'                 Optional Description_UnitTest As String = {"Description"} _
'        ) As State_UnitTest
''-----------------------------------------------------------------------------------------------------'
'
'    '``````````````````````````````````````````````````````````'
'    Dim Start_Time As Double, Status_UnitTest As State_UnitTest
'    Dim Need_RecompileAndRestart  As Boolean, UFrm   As Object
'    Dim Execution_Time     As Double, Show_Result    As Boolean
'    Static Flag_Recompile  As Boolean, Flag_InitTest As Boolean
'    Dim Flag_ShowUI As Boolean, Result_UnitTest As Variant
'    '``````````````````````````````````````````````````````````'
'
'    '`````````````````````````````````````````````````````````````````'
'    If UserForms.Count = 0 Then Exit Function Else Flag_ShowUI = False
'    '`````````````````````````````````````````````````````````````````'
'
'    '````````````````````````````````````````````````````````````````````````````'
'    For Each UFrm In UserForms
'        If TypeName(UFrm) = UI_vbFramework.Name Then Flag_ShowUI = True: Exit For
'    Next UFrm
'    '````````````````````````````````````````````````````````````````````````````'
'
'    '````````````````````````````````````'
'    If Not Flag_ShowUI Then Exit Function
'    '````````````````````````````````````'
'
'    '````````````````````````````````````````````````'
'    If Reset_StaticVariable Then
'        Flag_InitTest = False: Flag_Recompile = False
'        Exit Function
'    End If
'    '````````````````````````````````````````````````'
'
'    '``````````````````````````````````'
'    If Flag_InitTest Then Exit Function
'    '``````````````````````````````````'
'
'    '`````````````````````````````````````'
'    If Not Flag_Recompile Then
'        GoSub UnitTest_PreProcessing
'        GoSub UnitTest_RecompileAndRestart
'    End If
'    '`````````````````````````````````````'
'
'    '``````````````````````````````````````````````````````````````````'
'    Start_Time = Micro_Timer: GoSub UnitTest_Code: Flag_InitTest = True
'    Execution_Time = Round(Micro_Timer - Start_Time, 7&)
'    '``````````````````````````````````````````````````````````````````'
'
'    '`````````````````````````````````'
'    Call Fill_MainStructure( _
'              Name_Group, _
'              Description_UnitTest, _
'              Name_UTest, _
'              Execution_Time, _
'              Result_UnitTest, _
'              Status_UnitTest, _
'              Show_Result _
'         )
'
'    Exit Function
'    '`````````````````````````````````'
'
''```````````````````````````````````````````````````````````````````'
'UnitTest_Code: ' // Секция кода для проведения основного теста
'
'    Show_Result = {True}
'    Result_UnitTest = {Test_Function()}
'
'    If {True} Then
'        Status_UnitTest = Type_Success
'    ElseIf {False} Then
'        Status_UnitTest = Type_Ignored
'    Else
'        Status_UnitTest = Type_Failed
'    End If
'
'    Run_UnitTest_{GroupName}_{TestName} = Status_UnitTest
'
'    Return
''```````````````````````````````````````````````````````````````````'
'
''`````````````````````````````````````````````````````````````````````````````````````````````````````'
'UnitTest_PreProcessing: ' // Секция кода для предварительной обработки или подготовки данных для теста
'
'    ' // Код
'
'    Return
''`````````````````````````````````````````````````````````````````````````````````````````````````````'
'
''`````````````````````````````````````````````````````````````````````````````````````````````````````'
'UnitTest_RecompileAndRestart: ' // Секция кода для перекомпиляции проекта и повторного запуска теста!
'
'    Need_RecompileAndRestart = False
'
'    If (Not Flag_Recompile) And Need_RecompileAndRestart Then
'        Run_UnitTest_{GroupName}_{TestName} = Type_Recompile: Flag_Recompile = True
'        Deferred_ProcedureCall "Internal_Call": Exit Function
'    End If
'
'    Return
''`````````````````````````````````````````````````````````````````````````````````````````````````````'
'
''-----------------------------------------------------------------------------------------------------'
'End Function
''====================================================================================================='


'----------------------------------------------------------------------------------------------------------------------------'
'¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦'
'----------------------------------------------------------------------------------------------------------------------------'

' // <<< UT_Test's >>>

'======================================================================================================'
Private Function Run_UnitTest_BasisTypes_String_MD5( _
                 ByVal Inx_CurrentTest As Long, _
                 Optional Name_UTest As String = "-", _
                 Optional Reset_StaticVariable As Boolean = False, _
                 Optional Name_Group As String = "Базовые типы", _
                 Optional Description_UnitTest As String = "Тестирование хеш-функции MD5 на строках" _
        ) As State_UnitTest
'------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Start_Time As Double, Status_UnitTest As State_UnitTest
    Dim Need_RecompileAndRestart  As Boolean, UFrm   As Object
    Dim Execution_Time     As Double, Show_Result    As Boolean
    Static Flag_Recompile  As Boolean, Flag_InitTest As Boolean
    Dim Flag_ShowUI As Boolean, Result_UnitTest As Variant
    '``````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If UserForms.Count = 0 Then Exit Function Else Flag_ShowUI = False
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    For Each UFrm In UserForms
        If TypeName(UFrm) = UI_vbFramework.Name Then Flag_ShowUI = True: Exit For
    Next UFrm
    '````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    If Not Flag_ShowUI Then Exit Function
    '````````````````````````````````````'

    '````````````````````````````````````````````````'
    If Reset_StaticVariable Then
        Flag_InitTest = False: Flag_Recompile = False
        Exit Function
    End If
    '````````````````````````````````````````````````'

    '``````````````````````````````````'
    If Flag_InitTest Then Exit Function
    '``````````````````````````````````'

    '`````````````````````````````````````'
    If Not Flag_Recompile Then
        GoSub UnitTest_PreProcessing
        GoSub UnitTest_RecompileAndRestart
    End If
    '`````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Start_Time = Micro_Timer: GoSub UnitTest_Code: Flag_InitTest = True
    Execution_Time = Round(Micro_Timer - Start_Time, 7&)
    '``````````````````````````````````````````````````````````````````'

    '`````````````````````````````````'
    Call Fill_MainStructure( _
              Name_Group, _
              Description_UnitTest, _
              Name_UTest, _
              Execution_Time, _
              Result_UnitTest, _
              Status_UnitTest, _
              Show_Result _
         )

    Exit Function
    '`````````````````````````````````'

'````````````````````````````````````````````````````````````````````````'
UnitTest_Code: ' // Секция кода для проведения основного теста

    Show_Result = True
    Result_UnitTest = VBD_Kit_Hashing.Get_HashSumm_Data("_UTest", , MD_5)

    If Len(Result_UnitTest) > 0 Then
        Status_UnitTest = Type_Success
    ElseIf False Then
        Status_UnitTest = Type_Ignored
    Else
        Status_UnitTest = Type_Failed
    End If

    Run_UnitTest_BasisTypes_String_MD5 = Status_UnitTest

    Return
'````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_PreProcessing: ' // Секция кода для предварительной обработки или подготовки данных для теста

    ' // Код

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_RecompileAndRestart: ' // Секция кода для перекомпиляции проекта и повторного запуска теста!

    Need_RecompileAndRestart = False

    If (Not Flag_Recompile) And Need_RecompileAndRestart Then
        ' // Run_UnitTest_{GroupName}_{TestName} = Type_Recompile: Flag_Recompile = True
        Deferred_ProcedureCall "Internal_Call": Exit Function
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'======================================================================================================'
Private Function Run_UnitTest_BasisTypes_String_SHA1( _
                 ByVal Inx_CurrentTest As Long, _
                 Optional Name_UTest As String = "-", _
                 Optional Reset_StaticVariable As Boolean = False, _
                 Optional Name_Group As String = "Базовые типы", _
                 Optional Description_UnitTest As String = "Тестирование хеш-функции SHA1 на строках" _
        ) As State_UnitTest
'------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Start_Time As Double, Status_UnitTest As State_UnitTest
    Dim Need_RecompileAndRestart  As Boolean, UFrm   As Object
    Dim Execution_Time     As Double, Show_Result    As Boolean
    Static Flag_Recompile  As Boolean, Flag_InitTest As Boolean
    Dim Flag_ShowUI As Boolean, Result_UnitTest As Variant
    '``````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If UserForms.Count = 0 Then Exit Function Else Flag_ShowUI = False
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    For Each UFrm In UserForms
        If TypeName(UFrm) = UI_vbFramework.Name Then Flag_ShowUI = True: Exit For
    Next UFrm
    '````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    If Not Flag_ShowUI Then Exit Function
    '````````````````````````````````````'

    '````````````````````````````````````````````````'
    If Reset_StaticVariable Then
        Flag_InitTest = False: Flag_Recompile = False
        Exit Function
    End If
    '````````````````````````````````````````````````'

    '``````````````````````````````````'
    If Flag_InitTest Then Exit Function
    '``````````````````````````````````'

    '`````````````````````````````````````'
    If Not Flag_Recompile Then
        GoSub UnitTest_PreProcessing
        GoSub UnitTest_RecompileAndRestart
    End If
    '`````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Start_Time = Micro_Timer: GoSub UnitTest_Code: Flag_InitTest = True
    Execution_Time = Round(Micro_Timer - Start_Time, 7&)
    '``````````````````````````````````````````````````````````````````'

    '`````````````````````````````````'
    Call Fill_MainStructure( _
              Name_Group, _
              Description_UnitTest, _
              Name_UTest, _
              Execution_Time, _
              Result_UnitTest, _
              Status_UnitTest, _
              Show_Result _
         )

    Exit Function
    '`````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````'
UnitTest_Code: ' // Секция кода для проведения основного теста

    Show_Result = True
    Result_UnitTest = VBD_Kit_Hashing.Get_HashSumm_Data("_UTest", , SHA_1)

    If Len(Result_UnitTest) > 0 Then
        Status_UnitTest = Type_Success
    ElseIf False Then
        Status_UnitTest = Type_Ignored
    Else
        Status_UnitTest = Type_Failed
    End If

    Run_UnitTest_BasisTypes_String_SHA1 = Status_UnitTest

    Return
'`````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_PreProcessing: ' // Секция кода для предварительной обработки или подготовки данных для теста

    ' // Код

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_RecompileAndRestart: ' // Секция кода для перекомпиляции проекта и повторного запуска теста!

    Need_RecompileAndRestart = False

    If (Not Flag_Recompile) And Need_RecompileAndRestart Then
        ' // Run_UnitTest_{GroupName}_{TestName} = Type_Recompile: Flag_Recompile = True
        Deferred_ProcedureCall "Internal_Call": Exit Function
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'------------------------------------------------------------------------------------------------------'
End Function
'======================================================================================================'


'========================================================================================================'
Private Function Run_UnitTest_BasisTypes_String_SHA256( _
                 ByVal Inx_CurrentTest As Long, _
                 Optional Name_UTest As String = "-", _
                 Optional Reset_StaticVariable As Boolean = False, _
                 Optional Name_Group As String = "Базовые типы", _
                 Optional Description_UnitTest As String = "Тестирование хеш-функции SHA256 на строках" _
        ) As State_UnitTest
'--------------------------------------------------------------------------------------------------------'

    '``````````````````````````````````````````````````````````'
    Dim Start_Time As Double, Status_UnitTest As State_UnitTest
    Dim Need_RecompileAndRestart  As Boolean, UFrm   As Object
    Dim Execution_Time     As Double, Show_Result    As Boolean
    Static Flag_Recompile  As Boolean, Flag_InitTest As Boolean
    Dim Flag_ShowUI As Boolean, Result_UnitTest As Variant
    '``````````````````````````````````````````````````````````'

    '`````````````````````````````````````````````````````````````````'
    If UserForms.Count = 0 Then Exit Function Else Flag_ShowUI = False
    '`````````````````````````````````````````````````````````````````'

    '````````````````````````````````````````````````````````````````````````````'
    For Each UFrm In UserForms
        If TypeName(UFrm) = UI_vbFramework.Name Then Flag_ShowUI = True: Exit For
    Next UFrm
    '````````````````````````````````````````````````````````````````````````````'

    '````````````````````````````````````'
    If Not Flag_ShowUI Then Exit Function
    '````````````````````````````````````'

    '````````````````````````````````````````````````'
    If Reset_StaticVariable Then
        Flag_InitTest = False: Flag_Recompile = False
        Exit Function
    End If
    '````````````````````````````````````````````````'

    '``````````````````````````````````'
    If Flag_InitTest Then Exit Function
    '``````````````````````````````````'

    '`````````````````````````````````````'
    If Not Flag_Recompile Then
        GoSub UnitTest_PreProcessing
        GoSub UnitTest_RecompileAndRestart
    End If
    '`````````````````````````````````````'

    '``````````````````````````````````````````````````````````````````'
    Start_Time = Micro_Timer: GoSub UnitTest_Code: Flag_InitTest = True
    Execution_Time = Round(Micro_Timer - Start_Time, 7&)
    '``````````````````````````````````````````````````````````````````'

    '`````````````````````````````````'
    Call Fill_MainStructure( _
              Name_Group, _
              Description_UnitTest, _
              Name_UTest, _
              Execution_Time, _
              Result_UnitTest, _
              Status_UnitTest, _
              Show_Result _
         )

    Exit Function
    '`````````````````````````````````'

'````````````````````````````````````````````````````````````````'
UnitTest_Code: ' // Секция кода для проведения основного теста

    Show_Result = True
    Result_UnitTest = VBD_Kit_Hashing.Get_HashSumm_Data("_UTest")

    If Len(Result_UnitTest) > 0 Then
        Status_UnitTest = Type_Success
    ElseIf False Then
        Status_UnitTest = Type_Ignored
    Else
        Status_UnitTest = Type_Failed
    End If

    Run_UnitTest_BasisTypes_String_SHA256 = Status_UnitTest

    Return
'````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_PreProcessing: ' // Секция кода для предварительной обработки или подготовки данных для теста

    ' // Код

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'`````````````````````````````````````````````````````````````````````````````````````````````````````'
UnitTest_RecompileAndRestart: ' // Секция кода для перекомпиляции проекта и повторного запуска теста!

    Need_RecompileAndRestart = False

    If (Not Flag_Recompile) And Need_RecompileAndRestart Then
        ' // Run_UnitTest_{GroupName}_{TestName} = Type_Recompile: Flag_Recompile = True
        Deferred_ProcedureCall "Internal_Call": Exit Function
    End If

    Return
'`````````````````````````````````````````````````````````````````````````````````````````````````````'

'--------------------------------------------------------------------------------------------------------'
End Function
'========================================================================================================'


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
