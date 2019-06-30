Attribute VB_Name = "RRAES_Globals"
Option Explicit

'Declare SAP High Level Objects
Public oLogonControl As Object
Public oConnection As SAPLogonCtrl.Connection
Public oSAPFunctions As Object

' Declare SAP Functions
Public SAPCheckTransaction As SAPFunctionsOCX.Function
Public SAPReadPerson As SAPFunctionsOCX.Function
Public SAPReadOperation As SAPFunctionsOCX.Function
Public SAPReadOpText As SAPFunctionsOCX.Function
Public SAPReadOpPRTs As SAPFunctionsOCX.Function
Public SAPMakeBooking As SAPFunctionsOCX.Function
Public SAPMakePPBooking As SAPFunctionsOCX.Function
Public SAPCancelBooking As SAPFunctionsOCX.Function
Public SAPCheckPassword As SAPFunctionsOCX.Function
Public SAPShowBookings As SAPFunctionsOCX.Function
Public SAPFindMileStone As SAPFunctionsOCX.Function
Public SAPUpdateMileStone As SAPFunctionsOCX.Function
Public SAPReadPrevBookings As SAPFunctionsOCX.Function
Public SAPReadBoBo As SAPFunctionsOCX.Function
Public SAPMakeBoBo As SAPFunctionsOCX.Function
Public SAPSetPassword As SAPFunctionsOCX.Function
Public SAPReadInstallation As SAPFunctionsOCX.Function
Public SAPJumpToOP As SAPFunctionsOCX.Function
Public SAPUpdSFDCInstances As SAPFunctionsOCX.Function
Public SAPReadSysMessages As SAPFunctionsOCX.Function
Public SAPRetrieveRecValues As SAPFunctionsOCX.Function
Public SAPUpdateRecValues As SAPFunctionsOCX.Function
Public SAPUnlockBoBo As SAPFunctionsOCX.Function
Public SAPReadConfig As SAPFunctionsOCX.Function
Public SAPChkPrevOps As SAPFunctionsOCX.Function
Public SAPReadReasons As SAPFunctionsOCX.Function
Public SAPReadYieldScrap As SAPFunctionsOCX.Function
Public SAPReadPPSerNos As SAPFunctionsOCX.Function
Public SAPReadNextWrkCtr As SAPFunctionsOCX.Function
Public SAPUpdatePPSerNoLocn As SAPFunctionsOCX.Function

'Declare SAP Function Tables
Public DocumentsTable As SAPTableFactoryCtrl.Table
Public PRTsTable As SAPTableFactoryCtrl.Table
Public OpTextTable As SAPTableFactoryCtrl.Table
Public PrevBookingsTable As SAPTableFactoryCtrl.Table
Public BoBoTable As SAPTableFactoryCtrl.Table
Public SysMessagesTable As SAPTableFactoryCtrl.Table
Public RetrievedRecDateValuesTable As SAPTableFactoryCtrl.Table
Public RetrievedRecOrderValuesTable As SAPTableFactoryCtrl.Table
Public RecValuesToRetrieveTable As SAPTableFactoryCtrl.Table
Public PlantCfgTable As SAPTableFactoryCtrl.Table
Public OrderTypeCfgTable As SAPTableFactoryCtrl.Table
Public OpStatusTable As SAPTableFactoryCtrl.Table
Public OrderStatusTable As SAPTableFactoryCtrl.Table
Public PlantBldgCfgTable As SAPTableFactoryCtrl.Table
Public PackagesTable As SAPTableFactoryCtrl.Table
Public ReasonsTable As SAPTableFactoryCtrl.Table
Public CmdLineDocRefs As SAPTableFactoryCtrl.Table
Public MethodsTable As SAPTableFactoryCtrl.Table
Public MethodParmsTable As SAPTableFactoryCtrl.Table
Public DataCarriersTable As SAPTableFactoryCtrl.Table
Public PPSerNoTable As SAPTableFactoryCtrl.Table
Public PPSerNoLocnTable As SAPTableFactoryCtrl.Table
Public IntPPSerNoLocnTable As SAPTableFactoryCtrl.Table


'Object Array for Methods Processing
Public MethodObjs() As Variant
Public Method() As MethodInfo
Public NamedCells() As Variant
Public SAPGlobalVars As Collection

'Declare Public Constants
Public Const cnApplicationName = "SFDC"
'Public Const cnSFDCVersion = "3.0"
Public Const cnSAPTrue = "X"
Public Const cnSAPFalse = " "
Public Const cnBookOffAction = "O"
Public Const cnBookOnAction = "C"
Public Const cnBoBoWIPStatus = "W"
Public Const cnBookOnBookOff = "BOBO"
Public Const cnPostEventBooking = "PSTE"
Public Const cnWorkBook = "WB"
Public Const cnMultiBook = "MB"
Public Const cnZeroTimeConfirm = "ZTC"
Public Const cnRepairOrderType = "ZRPR"
Public Const cnNetworkOrderType = "PS04"
Public Const cnDiversionOrderType = "ZDIV"
Public Const cnTECO = "I0045"
Public Const cnCNF = "I0009"
Public Const cnCheckStartup = "S"
Public Const cnCheckRunning = "R"
Public Const cnCheckCloseDown = "C"
'Note the following constant is used as the key for encryption - DO NOT CHANGE IT
Public Const cnInstallationNotFound = "INSTALLATION_NOT_REGISTERED_IN_SAP"
Public Const cnChecked = 1
Public Const cnUnChecked = 0
Public Const cnYes = "Y"
Public Const cnNo = "N"
Public Const cnDialogTitleLogon = "SFDC"
Public Const cnDialogTitleCheckPass = "Check Password"
Public Const cnDialogTitleOpInfo = "Operation Information"
Public Const cnDialogTitleRecCntl = "Recording Controller"
Public Const cnDialogTitleWorkBook = "Work Booking"
Public Const cnTemporaryFileAddress = "C:\Temp\"
Public Const cnTemporaryFileName = "RecDoc.txt"
Public Const cnErrorFileNotFound = 2&
Public Const cnErrorPathNotFound = 3&
Public Const cnErrorBadFormat = 11&
Public Const cnSeErrorAccessDenied = 5            '  access denied
Public Const cnSeErrorAssocIncomplete = 27
Public Const cnSeErrorDDEBusy = 30
Public Const cnSeErrorDDEFail = 29
Public Const cnSeErrorDDETimeout = 28
Public Const cnSeErrorDLLNotFound = 32
Public Const cnSeErrorFNF = 2                     '  file not found
Public Const cnSeErrorNoAssoc = 31
Public Const cnSeErrorOOM = 8                     '  out of memory
Public Const cnSeErrorPNF = 3                     '  path not found
Public Const cnSeErrorShare = 26
Public Const cnLocaleIDate = &H21        '  short date format ordering
Public Const cnLocaleSDate = &H1D        '  date separator
Public Const cnHWND_TOPMOST = -1
'Public Const cnHWND_TOPMOST = 0         ' Stop forms coming to the top
Public Const cnSWP_NOMOVE = &H2
Public Const cnSWP_NOSIZE = &H1
Public Const cnERROR_SUCCESS = 0&
Public Const cnMAX_DOCUMENTS = 19
Public Const cnINVALID_HANDLE_VALUE = -1
Public Const cnMAX_PATH = 260
Public Const cnAPLHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!?%^&*()-_+=[]{};;@'~#,.<>/?\"
Public Const cnAPLHANUMERIC = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!?%^&*()-_+=[]{};;@'~#,.<>/?\0123456789"
Public Const cnNUMERIC = "0123456789"
Public Const cnExport = "E"
Public Const cnImport = "I"
Public Const cnInternal = "I"
Public Const cnExternal = "E"
Public Const cnStructure = "S"
Public Const cnTable = "T"
Public Const cnRefCellIndicator = "?"
Public Const cnSAPMethod = "SAP"
Public Const cnPosnOffset = 3
Public Const cnMax_UserName As Long = 15

'Declare LOCALE Constants
Public Const vbZLString As String = ""
  Public Const API_FALSE As Long = &H0&
  Public Const API_TRUE As Long = &H1&

  Public Const LOCALE_SYSTEM_DEFAULT As Long = &H800
  Public Const LOCALE_USER_DEFAULT As Long = &H400

  Public Const LOCALE_SLIST As Long = &HC         '  list item separator
  Public Const LOCALE_IMEASURE As Long = &HD         '  0 as long = metric, 1 as long = US
  
  Public Const LOCALE_SDECIMAL As Long = &HE         '  decimal separator
  Public Const LOCALE_STHOUSAND As Long = &HF         '  thousand separator
  Public Const LOCALE_SGROUPING As Long = &H10        '  digit grouping
  Public Const LOCALE_IDIGITS As Long = &H11        '  number of fractional digits
  Public Const LOCALE_ILZERO As Long = &H12        '  leading zeros for decimal
  Public Const LOCALE_SNATIVEDIGITS As Long = &H13        '  native ascii 0-9
  
  Public Const LOCALE_ICOUNTRY As Long = &H5        '  Country short ID
  Public Const LOCALE_SCOUNTRY As Long = &H6        '  Country long name
  
  Public Const LOCALE_SCURRENCY As Long = &H14        '  local monetary symbol
  Public Const LOCALE_SINTLSYMBOL As Long = &H15        '  intl monetary symbol
  Public Const LOCALE_SMONDECIMALSEP As Long = &H16        '  monetary decimal separator
  Public Const LOCALE_SMONTHOUSANDSEP As Long = &H17        '  monetary thousand separator
  Public Const LOCALE_SMONGROUPING As Long = &H18        '  monetary grouping
  Public Const LOCALE_ICURRDIGITS As Long = &H19        '  # local monetary digits
  Public Const LOCALE_IINTLCURRDIGITS As Long = &H1A        '  # intl monetary digits
  Public Const LOCALE_ICURRENCY As Long = &H1B        '  positive currency mode
  Public Const LOCALE_INEGCURR As Long = &H1C        '  negative currency mode
  
  Public Const LOCALE_SDATE As Long = &H1D        '  date separator
  Public Const LOCALE_STIME As Long = &H1E        '  time separator
  Public Const LOCALE_SSHORTDATE As Long = &H1F        '  short date format string
  Public Const LOCALE_SLONGDATE As Long = &H20        '  long date format string
  Public Const LOCALE_STIMEFORMAT As Long = &H1003      '  time format string
  Public Const LOCALE_IDATE As Long = &H21        '  short date format ordering
  Public Const LOCALE_ILDATE As Long = &H22        '  long date format ordering
  Public Const LOCALE_ITIME As Long = &H23        '  time format specifier
  Public Const LOCALE_ICENTURY As Long = &H24        '  century format specifier
  Public Const LOCALE_ITLZERO As Long = &H25        '  leading zeros in time field
  Public Const LOCALE_IDAYLZERO As Long = &H26        '  leading zeros in day field
  Public Const LOCALE_IMONLZERO As Long = &H27        '  leading zeros in month field
  Public Const LOCALE_S1159 As Long = &H28        '  AM designator
  Public Const LOCALE_S2359 As Long = &H29        '  PM designator
  
  Public Const LOCALE_SDAYNAME1 As Long = &H2A        '  long name for Monday
  Public Const LOCALE_SDAYNAME2 As Long = &H2B        '  long name for Tuesday
  Public Const LOCALE_SDAYNAME3 As Long = &H2C        '  long name for Wednesday
  Public Const LOCALE_SDAYNAME4 As Long = &H2D        '  long name for Thursday
  Public Const LOCALE_SDAYNAME5 As Long = &H2E        '  long name for Friday
  Public Const LOCALE_SDAYNAME6 As Long = &H2F        '  long name for Saturday
  Public Const LOCALE_SDAYNAME7 As Long = &H30        '  long name for Sunday
  Public Const LOCALE_SABBREVDAYNAME1 As Long = &H31        '  abbreviated name for Monday
  Public Const LOCALE_SABBREVDAYNAME2 As Long = &H32        '  abbreviated name for Tuesday
  Public Const LOCALE_SABBREVDAYNAME3 As Long = &H33        '  abbreviated name for Wednesday
  Public Const LOCALE_SABBREVDAYNAME4 As Long = &H34        '  abbreviated name for Thursday
  Public Const LOCALE_SABBREVDAYNAME5 As Long = &H35        '  abbreviated name for Friday
  Public Const LOCALE_SABBREVDAYNAME6 As Long = &H36        '  abbreviated name for Saturday
  Public Const LOCALE_SABBREVDAYNAME7 As Long = &H37        '  abbreviated name for Sunday
  Public Const LOCALE_SMONTHNAME1 As Long = &H38        '  long name for January
  Public Const LOCALE_SMONTHNAME2 As Long = &H39        '  long name for February
  Public Const LOCALE_SMONTHNAME3 As Long = &H3A        '  long name for March
  Public Const LOCALE_SMONTHNAME4 As Long = &H3B        '  long name for April
  Public Const LOCALE_SMONTHNAME5 As Long = &H3C        '  long name for May
  Public Const LOCALE_SMONTHNAME6 As Long = &H3D        '  long name for June
  Public Const LOCALE_SMONTHNAME7 As Long = &H3E        '  long name for July
  Public Const LOCALE_SMONTHNAME8 As Long = &H3F        '  long name for August
  Public Const LOCALE_SMONTHNAME9 As Long = &H40        '  long name for September
  Public Const LOCALE_SMONTHNAME10 As Long = &H41        '  long name for October
  Public Const LOCALE_SMONTHNAME11 As Long = &H42        '  long name for November
  Public Const LOCALE_SMONTHNAME12 As Long = &H43        '  long name for December
  Public Const LOCALE_SABBREVMONTHNAME1 As Long = &H44        '  abbreviated name for January
  Public Const LOCALE_SABBREVMONTHNAME2 As Long = &H45        '  abbreviated name for February
  Public Const LOCALE_SABBREVMONTHNAME3 As Long = &H46        '  abbreviated name for March
  Public Const LOCALE_SABBREVMONTHNAME4 As Long = &H47        '  abbreviated name for April
  Public Const LOCALE_SABBREVMONTHNAME5 As Long = &H48        '  abbreviated name for May
  Public Const LOCALE_SABBREVMONTHNAME6 As Long = &H49        '  abbreviated name for June
  Public Const LOCALE_SABBREVMONTHNAME7 As Long = &H4A        '  abbreviated name for July
  Public Const LOCALE_SABBREVMONTHNAME8 As Long = &H4B        '  abbreviated name for August
  Public Const LOCALE_SABBREVMONTHNAME9 As Long = &H4C        '  abbreviated name for September
  Public Const LOCALE_SABBREVMONTHNAME10 As Long = &H4D        '  abbreviated name for October
  Public Const LOCALE_SABBREVMONTHNAME11 As Long = &H4E        '  abbreviated name for November
  Public Const LOCALE_SABBREVMONTHNAME12 As Long = &H4F        '  abbreviated name for December
  Public Const LOCALE_SABBREVMONTHNAME13 As Long = &H100F
  
  Public Const LOCALE_SPOSITIVESIGN As Long = &H50        '  positive sign
  Public Const LOCALE_SNEGATIVESIGN As Long = &H51        '  negative sign
  Public Const LOCALE_IPOSSIGNPOSN As Long = &H52        '  positive sign position
  Public Const LOCALE_INEGSIGNPOSN As Long = &H53        '  negative sign position
  Public Const LOCALE_IPOSSYMPRECEDES As Long = &H54        '  mon sym precedes pos amt
  Public Const LOCALE_IPOSSEPBYSPACE As Long = &H55        '  mon sym sep by space from pos amt
  
  Public Const LOCALE_INEGSYMPRECEDES As Long = &H56        '  mon sym precedes neg amt
  Public Const LOCALE_INEGSEPBYSPACE As Long = &H57        '  mon sym sep by space from neg amt
  Public Const LOCALE_IDEFAULTANSICODEPAGE  As Long = &H1004   'def ansi code page
  
  Private Const IP_SUCCESS As Long = 0
  Private Const MAX_WSADescription As Long = 256
  Private Const MAX_WSASYSStatus As Long = 128
  Private Const WS_VERSION_REQD As Long = &H101
  Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
  Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
  Private Const MIN_SOCKETS_REQD As Long = 1
  Private Const SOCKET_ERROR As Long = -1
  Private Const ERROR_SUCCESS As Long = 0


'Declare Public User Defined Types (Structures)
Public Type SAPInfo
    Userid As String
    password As String
    Client As String * 3
    System As String * 3
    Server As String
    SystemNumber As Long
    GroupName As String
    RFCwithDialog As Integer
    LoadBalancing As String * 1
    TraceLevel As Long
    Language As String * 2
    SAPFormatDate As String * 8
    SAPFormatTime As String * 6
    DateTime As Date
    LocalDateTime As Date
    TimeOnEntry As Date
    ErrorMessage As String
    Debug As String * 1
    LogonErrorFile As String
    CodePage As String
    CodePageCheck As String * 1
End Type

Public Type OrderInfo
    Number As String
    Aufpl As String
    Desc As String
    Type As String
    Category As String
    Plant As String
    SalesOrder As String
    ControllingArea As String
    CreationDate As String
    Customer As String
    EngSerialNo As String
    ModuleNumber As String
    EngineType As String
    EngMark As String
    EngVar As String
    CorenaManId As String
    ATALocn As String
    RootFLocn As String
    RootNetWrk As String
    SOEqipment As String
    Operator As String
    ICAOCode As String
    CustAddr As String
    WorkScopeFileName As String
    DISLevel As String
    ExAircraftTailNo As String
    AircraftType As String
    Authority As String
    OrderQty As String
    OrderScrap As String
End Type

Public Type OpInfo
    Number As String
    Aplzl As String
    ConfNo As String
    Desc As String
    ConfirmStatus As String
    ActualStartDate As String
    Arbid As String
    WorkCentre As String
    Plant As String
    ActivityType As String
    WorkUnits As String
    LongTextExists As String * 1
    TLOpChangeDate As String
    PrevOpsCNF As String * 1
    Bookable As String * 1
    Yield As String
    Scrap As String
End Type

Public Type ConfInfo
    BackFlush As String * 1
    PostDate As String
    EndDate As String
    EndTime As String
    ActWork As String
    WorkCentre As String
    Plant As String
    WorkUnits As String
    ConfText As String
    DevReason As String
    ActType As String
    FinalConf As String * 1
    Complete As String * 1
    Yield As String
    Scrap As String
    Conf_no As String
End Type

Public Type PersonInfo
    ClockNumber As String
    PersName As String
    Plant As String
    PrevBookingsStartTime As String
    PrevBookingsEndTime As String
    password As String
    Deleted As String * 1
    CanWorkBook As String * 1
    PassWordInit As String * 1
    CICOStatus As String * 1
    CanMultiBook As String * 1
    CanRecordVals As String * 1
    CanFinalConf As String * 1
    CanSave As String * 1
End Type

Public Type PlantCfgInfo
    Plant As String
    BookingMode As String
    DZProc As String * 1
    WarnPeriod As Integer
    CancelPeriod As Integer
    AutoReg As String
    CanUpdRecDocs As String
    HelpFilePath As String
    TimeZone As String
    CheckString As String
    CheckStringPosn As String
    PersNoPosn As String
    CheckClockIn As String * 1
    FormTimeOut As Integer
    BookingTimeOut As Integer
    CardBufferLength As Integer
    ContactPhone As String
    URLRoot As String
    CorenaLinks As String * 1
    SOPath As String
    CancellationRestricted As String
End Type

Public Type OrderTypeCfgInfo
    Plant As String
    OrderType As String
    CanConfAfterFinal As String * 1
    CanConfAfterTECO As String * 1
    DefFinalConf As String * 1
    Chk4PrevCNF As String * 1
    ReasonsUpper As String
    ReasonsLower As String
End Type

Public Type MethodInfo
    Name As String
    FuncMod As String
    TabName As String
    InStruc As String
    OutStruc As String
    IntExt As String * 1
    SingleExe As String * 1
    ImpParmCnt As Integer
    ExpParmCnt As Integer
    StartParmRow As Integer
    MaxParms As Integer
    Executed As Boolean
End Type

Public Type MethodParmInfo
    MethodName As String
    Name As String
    ImpExp As String * 1
    Posn As String * 2
    Mandatory As String * 1
    VBName As String
    Excel As String * 1
    StrTab As String * 1
    DefVal As String
    NotImpt As String * 1
End Type

Public Type PlantBldgCfgInfo
    Plant As String
    Building As String
End Type

Public Type ComputerInfo
    Name As String
    Plant As String
    Building As String
    Location As String
    Stop As String * 1
    Disabled As String * 1
    SAPDebug As String * 1
    LocaleSystemDefault As Long
    ComOn As String * 1
    ComPort As Integer
    SFDCMulti As String * 1
End Type

Public Type BoBoRow
    OpNumber As String
    OpDesc As String
    ConfNo As String
    OrderNo As String
    WorkCentre As String
    OnDate As String
    OnTime As String
    Warning As String * 1
    OrderCategory As String
End Type

Public Type DocumentInfo
    Name As String
    DocType As String
    FileName As String
    ApplnType As String
    AppPath As String
    PrefixPath As String
    RegPath As String
    FullFilePath As String
    RevDate As String
End Type

Public Type MilestoneInfo
    Found As Boolean
    Aplzl As String
    Anlzu As String * 1
End Type

Public Type RecDocInfo
    SalesNo As String
    CellID As String
End Type

Public Type RecVals
    OrderNo  As String
    CellID As String
    RecValue As String
    SalesNo As String
    UserName As String
    UserNo As String
    Date As String
    Time As String
End Type

Public Type tagWorkBook
    Doc As Workbook
    Template As String
    Name As String
    Enabled As Boolean
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * cnMAX_PATH
    cAlternate As String * 14
End Type

Public Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Public Type PPSerNo
    OrderNo As String
    OpNumber As String
    PartNo As String
    SerialNo As String
    WIPWorkCentre As String
    ScrapInd As String
    UpdDate As String
    UpdTime As String
End Type

Public Type PPSerNoList
    PartNo As String
    SerialNo As String
End Type

Public Type PPSerNoLocn
    OrderNo As String
    OpNumber As String
    PartNo As String
    SerialNo As String
    WIPWorkCentre As String
    ScrapInd As String
    UpdDate As String
    UpdTime As String
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

'Declare Global Public UDT Variables
Public gblSAP As SAPInfo
Public gblOrder As OrderInfo
Public gblOperation As OpInfo
Public gblUser As PersonInfo
Public gblPlantCfg As PlantCfgInfo
Public gblOrdTypeCfg As OrderTypeCfgInfo
Public gblInstallation As ComputerInfo
Public gblMilestone As MilestoneInfo
Public gblDocument As DocumentInfo
Public gblRetrieval() As RecDocInfo
Public gblStored() As RecVals
Public gblChanged() As RecVals
Public gblRecordings() As RecVals
Public gblDocuments() As tagWorkBook
Public gblPlantBldgs() As PlantBldgCfgInfo


'Declare Global Public Variables
Public gblFunction As String
Public gblSFDCVersion As String
Public gblPersonValidated As Boolean
Public gblOperationWasFound As Boolean
Public gblBoBoCurrentRow As Integer
Public gblBatchTrackOK As Boolean
Public gblContinueBooking As Boolean
Public gblLastBookingTime As Date
Public gblExit_File As Boolean
Public gblDocCount As Integer
Public gblExcel As Application
Public gblbOpenDoc As Boolean
Public gblMethCellCnt As Integer
Public gblYieldScrapFail As Boolean
Public gblLocalDateFormat As String
Public gblLANUser As String
Public gblIPAddress As String
Public gblOPPlnt As String '* -- TCR7117 -- *
Public matnr As String     '* -- TCR7117 -- *
Public sernp As String     '* -- TCR7117 -- *
Public gdate As Date       '* -- TCR7117 -- *
Public PCNF As String      '* -- TCR7117 -- *
Public GCONF As String     '* -- TCR7117 -- *
Public GavailYield As Double '--TCR7117 --
Public gsval As String '--TCR7117 --


'Declare Win32API Functions
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Long, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
       Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
       ByVal lpKeyName As Any, ByVal lpString As Any, _
       ByVal lplFileName As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal _
    dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal _
    lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As _
    Long) As Long
      

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) _
    As Long

Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
    
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Declare Function GetUserName Lib "advapi32" _
   Alias "GetUserNameA" _
  (ByVal lpBuffer As String, _
   nSize As Long) As Long
      
Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long
  
  Private Declare Function gethostbyname Lib "wsock32.dll" _
  (ByVal hostname As String) As Long
  
Private Declare Function lstrlenA Lib "kernel32" _
  (lpString As Any) As Long

Private Declare Function WSAStartup Lib "wsock32.dll" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
   
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

Private Declare Function inet_ntoa Lib "wsock32.dll" _
  (ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, _
   ByVal Ptr As Long) As Long
                       
Private Declare Function gethostname Lib "wsock32.dll" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long

        
Public Function GetThreadUserName() As String

  'Retrieves the user name of the current
  'thread. This is the name of the user
  'currently logged onto the system. If
  'the current thread is impersonating
  'another client, GetUserName returns
  'the user name of the client that the
  'thread is impersonating.
   Dim buff As String
   Dim nSize As Long
   
   buff = Space$(cnMax_UserName)
   nSize = Len(buff)

   If GetUserName(buff, nSize) = 1 Then

      GetThreadUserName = TrimNull(buff)
      Exit Function

   End If

End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function



Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim success As Long
  
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
   
End Function


Public Sub SocketsCleanup()
  
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
   
End Sub
  

Public Function GetMachineName() As String

   Dim sHostName As String * 256
  
   If gethostname(sHostName, 256) = ERROR_SUCCESS Then
      GetMachineName = Trim$(sHostName)
   End If
  
End Function


Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address

   Dim nbytes As Long
   Dim ptrHosent As Long  'address of HOSENT structure
   Dim ptrName As Long    'address of name pointer
   Dim ptrAddress As Long 'address of address pointer
   Dim ptrIPAddress As Long
   Dim ptrIPAddress2 As Long

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset

     'Null-terminated list of addresses for the host.
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Note: Here we are retrieving
     'only the first address returned. To return more than
     'one, define sAddress as a string array and loop through
     'the 4-byte ptrIPAddress members returned. The last
     'item is a terminating null. All addresses are returned
     'in network byte order.
      ptrAddress = ptrHosent + 12
     
     'get the IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress2, ByVal ptrIPAddress, 4

      GetIPFromHostName = GetInetStrFromPtr(ptrIPAddress2)

   End If
  
End Function


Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
  
End Function


Private Function GetInetStrFromPtr(Address As Long) As String
 
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

Public Function SAPisConnected(SAP As SAPInfo, _
                             Computer As ComputerInfo, _
                             CheckType As String, _
                             BypassStop As Boolean) As Boolean

    Dim tmpTime, tmpDate As String
    Dim Dummy As Boolean
    
    'First test that we are still connected to SAP
    If oConnection.IsConnected <> tloRfcConnected Then
        'Try to reconnect to SAP
        If Connect2SAP = False Then
            SAPisConnected = False
            Exit Function
        Else
            SAPisConnected = True
        End If
    Else
        SAPisConnected = True
    End If
    
    'If still connected set Export Parameters for Call to Read Installation Info
    SAPReadInstallation.Exports("INST_IN").Value("HOST") = Computer.Name
    SAPReadInstallation.Exports("INST_IN").Value("APP") = cnApplicationName
    SAPReadInstallation.Exports("INST_IN").Value("APP_VER") = gblSFDCVersion
    SAPReadInstallation.Exports("CHECK_TYPE") = CheckType

    'Get the current Installation Details via an RFC call
    If SAPReadInstallation.Call = False Then
        MsgBox SAPReadInstallation.Exception, vbExclamation
        SAPisConnected = False
               
    Else
        SAP.System = SAPReadInstallation.Imports("SYSTEM_ID")
        'Set SAP Local Date & Time
        tmpTime = Format(SAPReadInstallation.Imports("LOCL_TIME"), "HH:MM:SS")
        tmpDate = Format(SAPReadInstallation.Imports("LOCL_DATE"), "YYYY/MM/DD")
        SAP.LocalDateTime = CDate(tmpDate + " " + tmpTime)
        'Set SAP Server Time
        SAP.SAPFormatTime = Format(SAPReadInstallation.Imports("SAP_TIME"), "HHMMSS")
        SAP.SAPFormatDate = Format(SAPReadInstallation.Imports("SAP_DATE"), "YYYYMMDD")
        tmpTime = Format(SAPReadInstallation.Imports("SAP_TIME"), "HH:MM:SS")
        tmpDate = Format(SAPReadInstallation.Imports("SAP_DATE"), "YYYY/MM/DD")
        SAP.DateTime = CDate(tmpDate + " " + tmpTime)
        
        Computer.Name = SAPReadInstallation.Imports("INST_OUT").Value("HOST")
        Computer.Plant = SAPReadInstallation.Imports("INST_OUT").Value("PLANT")
        Computer.Building = SAPReadInstallation.Imports("INST_OUT").Value("BUILDING")
        Computer.Location = SAPReadInstallation.Imports("INST_OUT").Value("LOCATION")
        Computer.Stop = SAPReadInstallation.Imports("INST_OUT").Value("STOP")
        Computer.Disabled = SAPReadInstallation.Imports("INST_OUT").Value("DISABLED")
        Computer.SAPDebug = SAPReadInstallation.Imports("INST_OUT").Value("SAP_DEBUG")
        Computer.ComOn = SAPReadInstallation.Imports("INST_OUT").Value("COM_ON")
        Computer.ComPort = SAPReadInstallation.Imports("INST_OUT").Value("COM_PORT")
        Computer.SFDCMulti = SAPReadInstallation.Imports("INST_OUT").Value("MULTI")
        
    End If
     
    'Do not perform close down checks if BypassStop is True
    If BypassStop = True Then GoTo Exit_SAPisConnected
        
    'Test for the installation Stop or Disabled Flag
    If Computer.Disabled = cnSAPTrue Then
      With frmStopMessage
          .lblDisableMsg.Caption = "SFDC HAS BEEN DISABLED ON THIS COMPUTER " & Chr(10) & _
                                    "PLEASE CONTACT THE SYSTEM ADMINISTRATOR"
          .Show vbModal
      End With
      'Perform Closedown for Installation before stopping application
      Dummy = SAPisConnected(gblSAP, gblInstallation, cnCheckCloseDown, True)
      End
    End If
      
    If Computer.Stop = cnSAPTrue Then
      With frmStopMessage
          .lblDisableMsg.Caption = "SFDC IS BEING STOPPED ON THIS COMPUTER" & Chr(10) & _
                                    "IT IS OK TO RESTART SFDC AFTERWARDS"
          .Show vbModal
      End With
      'Perform Closedown for Installation before stopping application
      Dummy = SAPisConnected(gblSAP, gblInstallation, cnCheckCloseDown, True)
      End
    End If

Exit_SAPisConnected:
  Exit Function
 
End Function

Public Function Connect2SAP() As Boolean

On Error GoTo file_error

    Connect2SAP = oConnection.Logon(frmLogon.hwnd, True)
        
    Select Case oConnection.IsConnected
        Case tloRfcConnected
            Connect2SAP = True
        Case Else
            Connect2SAP = False
            'write to the logon error file
            Open gblSAP.LogonErrorFile For Append Shared As #1
            Print #1, Now(); _
                Tab(25); RTrim(gblInstallation.Name); _
                Tab(41); gblSFDCVersion; _
                Tab(49); gblIPAddress; _
                Tab(69); gblLANUser; _
                Tab(84); gblSAP.System + gblSAP.Client; _
                Tab(91); gblSAP.Userid; _
                Tab(104); "Installation failed to connect to the identified client at the recorded (Local) time"
            
            Close #1

            'generate the error message for the user
            oConnection.LastError

    End Select
    
    Exit Function
    
file_error:

MsgBox "There was a problem recording the Logon Failure", vbExclamation
    
End Function



'From FAQ the Acrobat Reader can be found at
'"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\
'CurrentVersion\AppPaths\AcroRd32.exe"
'
'
Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)

Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long
Dim r As Long
Dim v As String

RetVal$ = ""
Const KEY_ALL_ACCESS As Long = &H3F
Const cnERROR_SUCCESS As Long = 0
Const REG_SZ As Long = 1

r = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)

If r <> cnERROR_SUCCESS Then GoTo Quit_Now

SZ = 256: v$ = String$(SZ, 0)
r = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)

If r = cnERROR_SUCCESS And dwType = REG_SZ Then
    SZ = SZ - 1
    RetVal$ = Left$(v$, SZ)
Else
    RetVal$ = "--Not String--"
End If

If hInKey = 0 Then r = RegCloseKey(hSubKey)

Quit_Now:
    RegGetString$ = RetVal$
    
End Function


Public Function EncryptPassword(password As String) As String
    Dim TempStr As String
    Dim i As Integer
    
    ' convert password to uppercase
    password = UCase(password)

    ' and then encrypt the password
    For i = 1 To Len(password)
        TempStr = TempStr + Hex(Asc(Mid(password, i, 1)) + 16 - i)
    Next i
    
    EncryptPassword = TempStr

End Function
Public Function DeCryptPassword(password As String) As String
    DeCryptPassword = Blowfish.DecryptString(password, cnInstallationNotFound)
End Function

Public Function GetHost() As String

    Dim TempStr As String * 100
    Dim StrLen As Integer
    Dim nSize As Long
    
    nSize = 20
    
    Call GetComputerName(TempStr, nSize)
    
    StrLen = InStr(TempStr, Chr(0))
    TempStr = Mid$(TempStr, 1, StrLen - 1)
    
    GetHost = TempStr
    
End Function

Public Sub GetSAPparameters()

    Dim SFDCIniFilePath As String
    Dim RedirectFilePath As String
    Dim ConnectionFilePath As String
    Dim TextLine As String
    Dim ServerName As String
    Dim FoundPosn As Integer
    Dim ParameterString As String
    Dim ValueString As String
    Dim File_Error_Code As Integer
    Dim SystemPath As String * 100
    Dim nSize As Long
    Dim ServerPath As String
    Dim RedirectOpenOK As Boolean
    
    
    'Setup error handler to trap non connection to LAN
    On Error GoTo error_handler
    
    ' Get System Path
    nSize = 20
    nSize = GetSystemDirectory(SystemPath, nSize)
    SFDCIniFilePath = Left(SystemPath, nSize)
    
    ' Set error code in case error returned when processing file
    File_Error_Code = 1 'Client-side ini file
    
    'Determine where the client-side .ini file is located and then......
    'open it
    SFDCIniFilePath = SFDCIniFilePath & "\sfdc.ini"
    Open SFDCIniFilePath For Input As #1 ' Open file.
    
    'Obtain paths for the server-side redirect files (hopefully, but not necessarily, on 3 different servers)
    'then try opening each one in turn, until successful.
    'nb. not only does this improve resilience, in the event of an individual server being down,
    'but also accommodates individual LAN userids with permissions only to view local servers
    
    File_Error_Code = 2 'Redirect File
    
    Do While Not EOF(1) ' Loop until end of client-side ini file (currently 3 lines, but could grow)
        Line Input #1, RedirectFilePath ' Read line into variable.
        'find the last backslash in the string and then use the position to....
        FoundPosn = InStrRev(RedirectFilePath, "\") - 1
        'truncate the string ready for concatenating a relative path for the connection details file
        ServerPath = Mid(RedirectFilePath, 1, FoundPosn)
        'Reset the file successful open flag for each iteration
        RedirectOpenOK = True
        'try to open the redirect file
        Open RedirectFilePath For Input As #2 ' Open Redirect file.
        If RedirectOpenOK Then Exit Do
    Loop
    
    Close #1
    
    'At this point test that at least one Redirect file was opened OK
    If Not RedirectOpenOK Then
        'Failed to open ANY of the 3 server-side redirect files so.....
        'Assume user is not logged onto network or doesn't have access authority
        MsgBox "Cannot read any Server-side files via the Network" & Chr(13) & Chr(13) _
        & "EITHER User does not have Access Authority OR is not Logged on to the Network" _
        , vbExclamation, cnDialogTitleLogon
        'stop processing at this point
         End
    End If
    
     
   
    'Process the server-side redirect ini file that was opened to check for a redirect
    'of this specific client (hostname) to a SAP instance other than the default
  
    Do While Not EOF(2) ' Loop until end of file.
        Line Input #2, TextLine ' Read line into variable.
        FoundPosn = InStr(1, TextLine, "=")
        If FoundPosn = 0 Then Exit Do
        ParameterString = UCase(Mid(TextLine, 1, FoundPosn - 1))
        ValueString = Mid(TextLine, FoundPosn + 1)
        Select Case ParameterString
            Case "DEFAULT"
                ConnectionFilePath = ValueString
            Case gblInstallation.Name
                ConnectionFilePath = ValueString
                Exit Do
        End Select
    Loop

    Close #2    ' Close Redirect file.
    
    'Test for an absolute or relative path to the ConnectionFile from the Redirect File
    'an absolute path will be preceded by "\\" anything else will be a relative path
    'we need to convert any relative path to an absolute path by concatenating it with
    'the first part of the original path used by the client-side ini file.
    '(adding a backslash between the two strings if necessary)
    'nb. this allows us to provide relative paths for the UK where we expect to copy the
    'set of server-side .ini files to three different servers without having to change them
    'in any way. Other non-UK sites can continue using absolute paths where only one set of
    'server-side .ini files is used.
    
    If Mid(ConnectionFilePath, 1, 2) <> "\\" Then
        If Mid(ConnectionFilePath, 1, 1) = "\" Then
            ConnectionFilePath = ServerPath + ConnectionFilePath
        Else
            ConnectionFilePath = ServerPath + "\" + ConnectionFilePath
        End If
    End If
    
    File_Error_Code = 3 'SAP Connection details file
    
    'Open and process the SAP connection details server-side ini file
    
    Open ConnectionFilePath For Input As #3
         
    Do While Not EOF(3) ' Loop until end of file.
        Line Input #3, TextLine ' Read line into variable.
        FoundPosn = InStr(1, TextLine, "=")
        If FoundPosn = 0 Then Exit Do
        ParameterString = UCase(Mid(TextLine, 1, FoundPosn - 1))
        ValueString = Mid(TextLine, FoundPosn + 1)
        Select Case ParameterString
            Case "USER"
                gblSAP.Userid = ValueString
            Case "PASSWORD"
                gblSAP.password = Left(DeCryptPassword(ValueString), 8)
            Case "CLIENT"
                gblSAP.Client = ValueString
            Case "SYSTEM"
                gblSAP.System = ValueString
            Case "SERVER"
                gblSAP.Server = ValueString
            Case "GROUPNAME"
                gblSAP.GroupName = ValueString
            Case "SYSTEM_NUMBER"
                gblSAP.SystemNumber = CLng(ValueString)
            Case "LANGUAGE"
                gblSAP.Language = ValueString
            Case "LOAD_BALANCING"
                gblSAP.LoadBalancing = ValueString
            Case "TRACE_LEVEL"
                gblSAP.TraceLevel = CLng(ValueString)
            Case "LOGON_ERROR_FILE"
                gblSAP.LogonErrorFile = ValueString
            Case "CODEPAGE"
                gblSAP.CodePage = ValueString
            Case "CODEPAGE_CHECKING"
                gblSAP.CodePageCheck = ValueString
                
        End Select
    Loop

    Close #3    ' Close parameters file.
    
    
    'Handle situation where the logon error file path is given as a relative path
    'in the SAP connection details file - the folder will NOT be with the ini files
    'because it needs universal write access. So, we need to go back another level in
    'the server path and assume that the error log folder will be on the same folder level as
    'the ini file folder.
    'nb. this is really a fallback 'cos we expect an absolute path pointing to one location
    
    'find the last backslash in the current serverpath string and then use the position to....
    FoundPosn = InStrRev(ServerPath, "\") - 1
    'truncate the string ready for concatenating a relative path for the connection details file
    ServerPath = Mid(ServerPath, 1, FoundPosn)
        
    'if it's specified as as an absolute path then leave it as-is
    If Mid(gblSAP.LogonErrorFile, 1, 2) <> "\\" Then
        If Mid(gblSAP.LogonErrorFile, 1, 1) = "\" Then
            gblSAP.LogonErrorFile = ServerPath + gblSAP.LogonErrorFile
        Else
            gblSAP.LogonErrorFile = ServerPath + "\" + gblSAP.LogonErrorFile
        End If
    End If
    
    
    Exit Sub
    
error_handler:

    Select Case File_Error_Code
    
        Case (1)
            'Display message to advise user that their SFDC.INI file is corrupted
            MsgBox "Shop Floor Data INI file is unreadable on this PC" & Chr(13) & Chr(13) _
            & "Please contact local IT support" _
            , vbExclamation, cnDialogTitleLogon
            ' stop processing at this point
            End
            
        Case (2)
            'Redirect file failed to open so flag and resume next
            RedirectOpenOK = False
            Resume Next
                      
            
        Case (3)
            'Connection details file failed to open
            FoundPosn = InStr(3, ConnectionFilePath, "\")
            ServerName = Mid(ConnectionFilePath, 3, FoundPosn - 3)
            'Assume user is not logged on to network so display message to advise user to do so
            MsgBox "Cannot Open Connection Details File on Server " & ServerName & Chr(13) & Chr(13) _
            & "EITHER User does not have Access Authority OR is not Logged On to the Network" _
            , vbExclamation, cnDialogTitleLogon
            'stop processing at this point
            End
            
        Case Else
            MsgBox Error$
            
    End Select
    
    
End Sub
Public Sub CreateSAPFunctionObjects()

   'Create the SAP Function Objects
   Set oSAPFunctions = CreateObject("SAP.Functions")
   oSAPFunctions.Connection = oConnection

   Set SAPReadInstallation = oSAPFunctions.Add("Z_SAP_CONNECT_TEST")
   Set SAPReadPerson = oSAPFunctions.Add("Z_READ_ZPERSONNEL")
   Set SAPReadOpText = oSAPFunctions.Add("Z_CONVERT_OPERATION_TEXT")
   Set SAPReadOperation = oSAPFunctions.Add("Z_READ_OPERATION_BY_CONFNO")
   Set SAPReadOpPRTs = oSAPFunctions.Add("Z_READ_OPERATION_PRTS")
   Set SAPCheckTransaction = oSAPFunctions.Add("Z_CHECK_CONFIRM_TRANSACTION")
   'Set SAPCheckPassword = oSAPFunctions.Add("ZRR_EMPLOYEE_CHECKPASSWORD")
   Set SAPMakeBooking = oSAPFunctions.Add("Z_MAKE_BOOKING")
   Set SAPMakePPBooking = oSAPFunctions.Add("Z_MAKE_PP_BOOKING")
   'Set SAPCancelBooking = oSAPFunctions.Add("BAPI_CONFIRMATION_CANCEL")
   Set SAPShowBookings = oSAPFunctions.Add("Z_RETRIEVE_BOOKINGS_BY_PERSON")
   Set SAPFindMileStone = oSAPFunctions.Add("Z_FIND_FOLLOWING_MILESTONE")
   Set SAPUpdateMileStone = oSAPFunctions.Add("Z_UPDATE_MILESTONE")
   Set SAPReadPrevBookings = oSAPFunctions.Add("Z_RETRIEVE_BOOKINGS_BY_PERSON")
   Set SAPSetPassword = oSAPFunctions.Add("Z_SET_PASSWORD")
   Set SAPCancelBooking = oSAPFunctions.Add("Z_CANCEL_BOOKING")
   Set SAPJumpToOP = oSAPFunctions.Add("Z_JUMP_TO_OPERATION")
   Set SAPUpdSFDCInstances = oSAPFunctions.Add("Z_RECORD_SFDC_VERSION")
   Set SAPReadSysMessages = oSAPFunctions.Add("Z_READ_SYSTEM_MESSAGES")
   Set SAPRetrieveRecValues = oSAPFunctions.Add("Z_RECORDING_RETRIEVE")
   Set SAPUpdateRecValues = oSAPFunctions.Add("Z_RECORDING_UPDATE")
   Set SAPReadBoBo = oSAPFunctions.Add("Z_READ_ZBOBO")
   Set SAPMakeBoBo = oSAPFunctions.Add("Z_BOOK_ON_OFF")
   Set SAPUnlockBoBo = oSAPFunctions.Add("Z_UNLOCK_ZBOBO")
   Set SAPReadConfig = oSAPFunctions.Add("Z_READ_SFDC_CFG")
   Set SAPChkPrevOps = oSAPFunctions.Add("Z_CHECK4_PREV_CNF_OPS")
   Set SAPReadReasons = oSAPFunctions.Add("Z_READ_VARIANCE_REASONS")
   Set SAPReadYieldScrap = oSAPFunctions.Add("ZREAD_PP_YIELDSCRAP")
   Set SAPReadPPSerNos = oSAPFunctions.Add("ZRETRIEVE_PP_SERNO")
   Set SAPReadNextWrkCtr = oSAPFunctions.Add("ZREAD_NEXT_WRKCTR")
   Set SAPUpdatePPSerNoLocn = oSAPFunctions.Add("Z_UPDATE_PP_SERNO_LOCN")

   'Create Object variables for the SAP Tables
   Set OpStatusTable = SAPReadOperation.Tables.Item("OP_STATUS")
   Set OrderStatusTable = SAPReadOperation.Tables.Item("ORDER_STATUS")
   Set PackagesTable = SAPReadOperation.Tables.Item("PACKAGES")
   Set PlantCfgTable = SAPReadConfig.Tables.Item("CFG1")
   Set OrderTypeCfgTable = SAPReadConfig.Tables.Item("CFG2")
   Set PlantBldgCfgTable = SAPReadConfig.Tables.Item("CFG3")
   Set DataCarriersTable = SAPReadConfig.Tables.Item("O_TDWN")
   Set MethodsTable = SAPReadConfig.Tables.Item("METHOD")
   Set MethodParmsTable = SAPReadConfig.Tables.Item("METHOD_PARM")
   Set ReasonsTable = SAPReadReasons.Tables.Item("REASONS")
   Set OpTextTable = SAPReadOpText.Tables.Item("OP_TEXT")
   Set DocumentsTable = SAPReadOpPRTs.Tables.Item("DOCUMENTS")
   Set PRTsTable = SAPReadOpPRTs.Tables.Item("TOOLS")
   Set PrevBookingsTable = SAPReadPrevBookings.Tables.Item("O_BOOKINGS")
   Set RetrievedRecOrderValuesTable = SAPRetrieveRecValues.Tables.Item("REC_ORDER_RETRIEVE")
   Set RetrievedRecDateValuesTable = SAPRetrieveRecValues.Tables.Item("REC_DATE_RETRIEVE")
   Set SysMessagesTable = SAPReadSysMessages.Tables.Item("SYS_MESSAGE")
   Set BoBoTable = SAPReadBoBo.Tables.Item("BOBO")
   Set CmdLineDocRefs = SAPReadOpText.Tables.Item("DOCREFS")
   Set PPSerNoTable = SAPReadPPSerNos.Tables.Item("PPSERNOS")
   Set PPSerNoLocnTable = SAPUpdatePPSerNoLocn.Tables.Item("SERNOLOCN")
   Set IntPPSerNoLocnTable = SAPUpdatePPSerNoLocn.Tables.Item("SERNOLOCN")


End Sub
Public Function ConvertPath(OrigPath As String, NewFileName As String) As String
    Dim PeriodPosn As Integer
    Dim LastSlashPosn As Integer
    Dim FileType As String
    Dim Path As String
    Dim i As Integer
    
    ' Determine Period Position
    PeriodPosn = InStr(OrigPath, ".")
    ' Determine Last Back Slash Position
    For i = (PeriodPosn - 1) To 1 Step -1
        If Mid(OrigPath, i, 1) = "\" Then
            LastSlashPosn = i
            Exit For
        End If
    Next i
    
    FileType = Mid(OrigPath, PeriodPosn)
    Path = Mid(OrigPath, 1, LastSlashPosn)
    
    ConvertPath = Path + NewFileName + FileType
    
End Function
Public Function ShellExErrMsg(ErrNumber As Variant) As String
     
    Select Case ErrNumber
      Case 0
        ShellExErrMsg = "The operating system is out of memory or resources"
      Case cnErrorFileNotFound
        ShellExErrMsg = "The specified file was not found"
      Case cnErrorPathNotFound
        ShellExErrMsg = "The specified path was not found"
      Case cnErrorBadFormat
        ShellExErrMsg = "The .EXE file is invalid (non Win32 .EXE or error in .EXE image)"
      Case cnSeErrorAccessDenied
        ShellExErrMsg = "The operating system denied access to the specified file"
      Case cnSeErrorAssocIncomplete
        ShellExErrMsg = "The filename association is incomplete or invalid"
      Case cnSeErrorDDEBusy
        ShellExErrMsg = "The DDE transaction could not be completed because other DDE transactions were being processed"
      Case cnSeErrorDDEFail
        ShellExErrMsg = "The DDE transaction failed"
      Case cnSeErrorDDETimeout
        ShellExErrMsg = "The DDE transaction could not be completed because the request timed out"
      Case cnSeErrorDLLNotFound
        ShellExErrMsg = "The specified dynamic-link library was not found"
      Case cnSeErrorFNF
        ShellExErrMsg = "The specified file was not found"
      Case cnSeErrorNoAssoc
        ShellExErrMsg = "There is no application associated with the given filename extension"
      Case cnSeErrorOOM
        ShellExErrMsg = "There was not enough memory to complete the operation"
      Case cnSeErrorPNF
        ShellExErrMsg = "The specified path was not found"
      Case cnSeErrorShare
        ShellExErrMsg = "A sharing violation occurred"
      Case Else
        ShellExErrMsg = ""
    End Select
    
    
End Function


Public Sub ResetLogonForm()

        gblOperationWasFound = False
        
        ' Set logon screen fields to null ready for new booking
        With frmLogon
            .ConfirmationNumber = ""
            .OrderNumber = ""
            .OpNumber = ""
            .WorkCentre = ""
            .OpDescription = ""
            .OrderDesc = ""
        End With


End Sub

Public Sub ReadSysMessages()

    Dim RowCount As Integer
    Dim i As Integer
    Dim MessageLine As String * 62
   
    
    SysMessagesTable.FreeTable 'remove existing data in table
   
   'then call the RFC
   If SAPReadSysMessages.Call = False Then
        'hide the list box 'cos there are no messages to display
        frmLogon.lstSysMessages.Visible = False
   Else
        'Clear the List Box
         frmLogon.lstSysMessages.Clear
        'determine no of rows
         RowCount = SysMessagesTable.RowCount
        'populate the list box
        For i = 1 To RowCount
            MessageLine = SysMessagesTable.Value(i, "EMTEXT")
            frmLogon.lstSysMessages.AddItem (MessageLine)
        Next
        'and display it
        frmLogon.lstSysMessages.Visible = True
        
   End If

End Sub


Public Function IsCodePageOK() As Boolean


    Dim LCID As Long
    Dim CurrentCodePage As String
    
    'Set default return value of RegionalSettingsOK
    IsCodePageOK = True
    
   
  'Check the Codepage is 1252, if not abort with error message
  
  'American National Standards Institute (ANSI) code page
  'associated with this locale. If the locale does not use
  'an ANSI code page, the value is 0. The maximum characters
  'allowed is six.
   LCID = GetSystemDefaultLCID()
   
   CurrentCodePage = GetUserLocaleInfo(LCID, LOCALE_IDEFAULTANSICODEPAGE)
   
   If CurrentCodePage <> gblSAP.CodePage Then
   
    IsCodePageOK = False
    MsgBox "Invalid Code Page on this machine - SFDC cannot be started", vbExclamation
    Exit Function
    
   End If
    

    
        
End Function

Public Sub SetLocalDateFormat()

    Dim x As Integer
    Dim lpLCData As String * 500
    Dim cchData As Long
    Dim oGetFormats As cGetLocalFormats
    
    'Instanciate the object (from the cGetLocalFormats class) and return the local system date format settings
    Set oGetFormats = New cGetLocalFormats
    
    With oGetFormats
        gblLocalDateFormat = .FourDigitYearDateFormat
    End With
    
    Set oGetFormats = Nothing
    
    cchData = 100

End Sub

Public Function ReadPerson(Person As PersonInfo) As Boolean

    Dim StartTime As String
    Dim EndTime As String
    Dim tmpDate As String
    Dim tmpDateTime As String * 18

    'Validate Check Number
    SAPReadPerson.Exports("PERSNO") = Person.ClockNumber
    SAPReadPerson.Exports("OP_PLNT") = gblOPPlnt ' -- TCR7117 --
    If SAPReadPerson.Call = False Then
      MsgBox SAPReadPerson.Exception, vbExclamation, cnDialogTitleCheckPass
      ReadPerson = False
      Exit Function
    Else
          
      With Person
        .Plant = SAPReadPerson.Imports("PERS_DETAILS").Value("ZPERS_PLNT")
        .password = SAPReadPerson.Imports("PERS_DETAILS").Value("PASSWORD")
        .PassWordInit = SAPReadPerson.Imports("PERS_DETAILS").Value("PASS_INIT")
        .PersName = SAPReadPerson.Imports("PERS_DETAILS").Value("ZPERS_NAME")
        .CICOStatus = SAPReadPerson.Imports("PERS_DETAILS").Value("CICO_STATUS")
        .CanRecordVals = SAPReadPerson.Imports("PERS_DETAILS").Value("REC_VALS")
        .CanMultiBook = SAPReadPerson.Imports("PERS_DETAILS").Value("MULTIPLE")
        .PrevBookingsStartTime = SAPReadPerson.Imports("PERS_DETAILS").Value("START_TIME")
        .PrevBookingsEndTime = SAPReadPerson.Imports("PERS_DETAILS").Value("END_TIME")
        .CanFinalConf = SAPReadPerson.Imports("PERS_DETAILS").Value("CANFIN")
        .CanSave = SAPReadPerson.Imports("PERS_DETAILS").Value("CANSAVE")
        .CanWorkBook = SAPReadPerson.Imports("PERS_DETAILS").Value("CANWB")
      End With
            
      
      ReadPerson = True
      
    End If

End Function

Public Function ReadOrderOp(Order As OrderInfo, _
                            Op As OpInfo, _
                            OrdTypeCfg As OrderTypeCfgInfo) As Boolean

    Dim ThisOp As OpInfo

   'Set export parameter for SAPReadOperation
   SAPReadOperation.Exports("CONFIRMATION_NO") = Op.ConfNo
   SAPReadOperation.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
   GCONF = Op.ConfNo ' -- TCR7117 --
   'Clear the Tables prior to the call
   OrderStatusTable.FreeTable
   OpStatusTable.FreeTable
   PackagesTable.FreeTable
   
   'Set Mouse Pointer to Hourglass during processing
   Screen.MousePointer = vbHourglass
   
   'then call the function module
   If SAPReadOperation.Call = False Then
        'set pointer to standard
        Screen.MousePointer = vbDefault
        MsgBox SAPReadOperation.Exception, vbExclamation, cnDialogTitleLogon
        ReadOrderOp = False
        Exit Function
   Else
        ReadOrderOp = True
        Screen.MousePointer = vbDefault
        With Order
            .Number = SAPReadOperation.Imports("ORDER").Value("AUFNR")
            .Desc = SAPReadOperation.Imports("ORDER").Value("KTEXT")
            .Aufpl = SAPReadOperation.Imports("ORDER").Value("AUFPL")
            .Plant = SAPReadOperation.Imports("ORDER").Value("WERKS")
            .Category = SAPReadOperation.Imports("ORDER").Value("AUTYP")
            .Type = SAPReadOperation.Imports("ORDER").Value("AUART")
            .ControllingArea = SAPReadOperation.Imports("ORDER").Value("KOKRS")
            .CreationDate = SAPReadOperation.Imports("ORDER").Value("ERDAT")
            .ATALocn = SAPReadOperation.Imports("ORDER").Value("ATA_LOCN")
            .WorkScopeFileName = SAPReadOperation.Imports("ORDER").Value("WS_FILENAME")
            .SalesOrder = SAPReadOperation.Imports("SO_INFO").Value("KDAUF")
            .Customer = SAPReadOperation.Imports("SO_INFO").Value("CUST_NAME")
            .EngSerialNo = SAPReadOperation.Imports("SO_INFO").Value("SERIAL_NO")
            .ModuleNumber = SAPReadOperation.Imports("SO_INFO").Value("MODULENO")
            .EngineType = SAPReadOperation.Imports("SO_INFO").Value("ENGTY")
            .EngMark = SAPReadOperation.Imports("SO_INFO").Value("ENGMK")
            .EngVar = SAPReadOperation.Imports("SO_INFO").Value("ENGVR")
            .CorenaManId = SAPReadOperation.Imports("SO_INFO").Value("MANID")
            .RootFLocn = SAPReadOperation.Imports("SO_INFO").Value("ROOT_FLOCN")
            .RootNetWrk = SAPReadOperation.Imports("SO_INFO").Value("ROOT_NETWK")
            .SOEqipment = SAPReadOperation.Imports("SO_INFO").Value("SO_EQUIP")
            .Operator = SAPReadOperation.Imports("SO_INFO").Value("OPERATOR")
            .ICAOCode = SAPReadOperation.Imports("SO_INFO").Value("ICAO_CODE")
            .DISLevel = SAPReadOperation.Imports("SO_INFO").Value("DISLEV")
            .AircraftType = SAPReadOperation.Imports("SO_INFO").Value("ACTYP")
            .ExAircraftTailNo = SAPReadOperation.Imports("SO_INFO").Value("EXATN")
            .Authority = SAPReadOperation.Imports("SO_INFO").Value("AUTHY")
        End With
        
        With Op
            .WorkCentre = SAPReadOperation.Imports("OPERATION").Value("WORKCENTRE")
            .Number = SAPReadOperation.Imports("OPERATION").Value("VORNR")
            .Desc = SAPReadOperation.Imports("OPERATION").Value("LTXA1")
            .Aplzl = SAPReadOperation.Imports("OPERATION").Value("APLZL")
            .LongTextExists = SAPReadOperation.Imports("OPERATION").Value("TXTSP")
            .ConfirmStatus = SAPReadOperation.Imports("OPERATION").Value("CONFIRM_STATUS")
            .Plant = SAPReadOperation.Imports("OPERATION").Value("WERKS")
            .ActivityType = SAPReadOperation.Imports("OPERATION").Value("LARNT")
            .Arbid = SAPReadOperation.Imports("OPERATION").Value("ARBID")
            .ActualStartDate = SAPReadOperation.Imports("OPERATION").Value("ISDD")
            .WorkUnits = SAPReadOperation.Imports("OPERATION").Value("ARBEH")
            .TLOpChangeDate = SAPReadOperation.Imports("TASK_LIST_OP").Value("AEDAT")
            .PrevOpsCNF = SAPReadOperation.Imports("OPERATION").Value("OK2CNF")
            .Bookable = SAPReadOperation.Imports("OPERATION").Value("BOOKABLE")
        End With
        gblOPPlnt = Op.Plant ' -- TCR7117 --
        
        'If PP Order set work units from VGE03 field
        If Order.Category = 10 Then
            Op.WorkUnits = SAPReadOperation.Imports("OPERATION").Value("VGE03")
        End If
        
        'Check for changes to Task List Op within warning period
        Call Check4OpChanges(Op)
        
        'Set OrderTypeCfg UDT for this Operation
        OrdTypeCfg.Plant = Op.Plant
        OrdTypeCfg.OrderType = Order.Type
        If SetOrdTypeConfig(OrdTypeCfg) = False Then
            MsgBox "CONFIGURATION DATA MISSING FOR" & Chr(10) & _
            "PLANT " & Op.Plant & " AND ORDER TYPE " & Order.Type, vbExclamation
            ReadOrderOp = False
            Exit Function
        End If
                    
        'Set Activity Type as appropriate where missing
        If Op.ActivityType = "" Then
            Select Case Order.Type
                Case cnRepairOrderType
                    Op.ActivityType = "T&M"
               Case cnNetworkOrderType
                    Op.ActivityType = "FIXED"
            End Select
        End If


   End If

End Function

Public Function SetPlantConfig(Cfg1 As PlantCfgInfo) As Boolean

   Dim Row As Integer, i As Integer

   'Initialise return value as False
   SetPlantConfig = False

   'Loop thru Plant Config Table until a Plant match is found
   'then read the row into the appropriate values of UDT
   For Row = 1 To PlantCfgTable.RowCount
      If PlantCfgTable.Value(Row, "PLANT") = Cfg1.Plant Then
         With Cfg1
            .BookingMode = PlantCfgTable.Value(Row, "BOOK_MODE")
            .AutoReg = PlantCfgTable.Value(Row, "AUTO_REG")
            .CancelPeriod = PlantCfgTable.Value(Row, "CANCEL")
            .CheckClockIn = PlantCfgTable.Value(Row, "TST_CICO")
            .DZProc = PlantCfgTable.Value(Row, "DZ_PROC")
            .HelpFilePath = PlantCfgTable.Value(Row, "HELP_FILE")
            .PersNoPosn = PlantCfgTable.Value(Row, "PERSNOPOS")
            .CheckString = PlantCfgTable.Value(Row, "CHKSTR")
            .CheckStringPosn = PlantCfgTable.Value(Row, "CHKSTRPOS")
            .CanUpdRecDocs = PlantCfgTable.Value(Row, "RECUPD")
            .TimeZone = PlantCfgTable.Value(Row, "TIME_ZONE")
            .WarnPeriod = PlantCfgTable.Value(Row, "WARN")
            .FormTimeOut = PlantCfgTable.Value(Row, "FORMTO")
            .BookingTimeOut = PlantCfgTable.Value(Row, "BOOKTO")
            .CardBufferLength = PlantCfgTable.Value(Row, "CARDBUFFER")
            .ContactPhone = PlantCfgTable.Value(Row, "CONTACTNO")
            .URLRoot = PlantCfgTable.Value(Row, "URL_ROOT")
            .CorenaLinks = PlantCfgTable.Value(Row, "CORENA_LNKS")
            .SOPath = PlantCfgTable.Value(Row, "SO_PATH")
            .CancellationRestricted = PlantCfgTable.Value(Row, "RESTRICT_CAN")
         End With

         For i = 1 To DataCarriersTable.RowCount
            If DataCarriersTable.Value(i, "DTTRG") = Cfg1.SOPath Then
               Cfg1.SOPath = DataCarriersTable.Value(i, "PRFXP")
               Exit For
            End If
         Next i

         SetPlantConfig = True
         Exit For
      End If
   Next Row

End Function
Public Function SetOrdTypeConfig(Cfg2 As OrderTypeCfgInfo) As Boolean

    Dim Row As Integer
    
    'Initialise return value as False
    SetOrdTypeConfig = False

    'Loop thru Plant Config Table until a Plant match is found
    'then read the row into the appropriate values of UDT
    For Row = 1 To OrderTypeCfgTable.RowCount
        If OrderTypeCfgTable.Value(Row, "PLANT") = Cfg2.Plant And _
           OrderTypeCfgTable.Value(Row, "ORD_TYP") = Cfg2.OrderType Then
            With Cfg2
                .CanConfAfterFinal = OrderTypeCfgTable.Value(Row, "FCAFC")
                .CanConfAfterTECO = OrderTypeCfgTable.Value(Row, "CATECO")
                .DefFinalConf = OrderTypeCfgTable.Value(Row, "DFC")
                .Chk4PrevCNF = OrderTypeCfgTable.Value(Row, "CHK4PREVCNF")
                .ReasonsLower = OrderTypeCfgTable.Value(Row, "REAS_LOWER")
                .ReasonsUpper = OrderTypeCfgTable.Value(Row, "REAS_UPPER")
            End With
            SetOrdTypeConfig = True
            Exit For
        End If
    Next Row
    
End Function

Public Function FindNextMilestone() As Boolean

        'Set export parameters for SAPFindMilestone
        SAPFindMileStone.Exports("I_AUFPL") = gblOrder.Aufpl
        SAPFindMileStone.Exports("I_VORNR") = gblOperation.Number
        
        'Call the function
        If SAPFindMileStone.Call = True Then
            'set the Milestone variables
            gblMilestone.Anlzu = True
            gblMilestone.Aplzl = SAPFindMileStone.Imports("O_APLZL")
            gblMilestone.Anlzu = SAPFindMileStone.Imports("O_ANLZU")
        Else
            gblMilestone.Anlzu = False
        End If

End Function

Public Sub Check4OpChanges(Op As OpInfo)

    Dim Threshold As Date

    Threshold = Date - gblPlantCfg.WarnPeriod
      
    If Op.TLOpChangeDate >= Threshold Then
        'And OrderCreationDate >= TLOpChangeDate Then
        MsgBox "WARNING - THIS OPERATION WAS CHANGED ON " & _
                Op.TLOpChangeDate & Chr(10) & _
                "PLEASE CHECK THE DOCUMENTATION CAREFULLY", _
                vbExclamation, cnDialogTitleLogon
    End If
        

End Sub


Public Function MakeConfirmation(Conf As ConfInfo, _
                                 Order As OrderInfo, _
                                 Op As OpInfo) As Boolean
                                 
    Dim BookingOK As String * 1
    Dim SAPErrorMessage As String
    Dim tmpSerNoLocnTable As SAPTableFactoryCtrl.Table
    
    'Set pointer to hourglass during SAP processing
    Screen.MousePointer = vbHourglass

    'Check order category to determine which booking function module to call
    If Order.Category = "10" Then
     
   'Don't send the yield/Scrap Qty and Final confirmation  while time booking ...
   If gblFunction <> "ZTC" Then
    If frmBookOff.optFullYield <> True Then
     If Val(gsval) <> Val(Conf.Yield) + Val(Conf.Scrap) Then
     Conf.FinalConf = ""
     Else
     If frmBookOff.optTimeOnly.Value = False Then '-- TCR7117 --
     Conf.FinalConf = "X"
     End If
     End If
     frmBatchTrack.txtYieldQty = ""          '-- TCR7117 --
     frmBatchTrack.txtScrapQty = ""          '-- TCR7117 --
     If frmBookOff.optTimeOnly.Value = True Then '-- TCR7117 --
     Conf.Yield = ""
     Conf.Scrap = ""
     End If
     Else
     If Val(gsval) = Val(Conf.Yield) + Val(Conf.Scrap) Then
      Conf.FinalConf = "X"
     End If
    End If
    gsval = ""
    Else
     If frmBatchTrack.optFullYield <> True Then
      If Val(gsval) <> Val(Conf.Yield) + Val(Conf.Scrap) Then
      Conf.FinalConf = ""
      End If
      Else
      If Val(gsval) = Val(Conf.Yield) + Val(Conf.Scrap) Then
      Conf.FinalConf = "X"
      End If
     End If
    End If '-- TCR7117 --
   gsval = ""
    'Check previous milestone Op is not completed then, don't send yield / scarp / final confirmation ..
    If PCNF = cnSAPFalse Then '-- TCR7117 --
     Conf.FinalConf = ""
     Conf.Yield = ""
     Conf.Scrap = ""
    End If
    PCNF = cnSAPTrue '-- TCR7117 --
    GavailYield = 0  '-- TCR7117 --
      'Set the parameters for the RFC
        With SAPMakePPBooking
            .Exports("APPLICATION") = Order.Category
            .Exports("I_AUFPL") = Order.Aufpl
            .Exports("I_APLZL") = Op.Aplzl
            .Exports("SAP_DEBUG") = gblInstallation.SAPDebug
            .Exports("YIELD") = Conf.Yield
            .Exports("SCRAP") = Conf.Scrap
        End With
        
   'Set export parameters for SAPMakePPBooking

        With SAPMakePPBooking.Exports("CONF_DETAILS")
            .Value("ORDER") = Order.Number
            .Value("ACTIVITY") = Op.Number
            .Value("PERS_NO") = gblUser.ClockNumber
            .Value("WORK_CNTR") = Conf.WorkCentre
            .Value("PLANT") = Conf.Plant
            .Value("POSTG_DATE") = Conf.PostDate
            .Value("END_DATE") = Conf.EndDate
            .Value("END_TIME") = Conf.EndTime
            .Value("ACT_WORK") = Conf.ActWork
            .Value("UN_WORK") = Conf.WorkUnits
            .Value("CONF_TEXT") = Conf.ConfText
            .Value("DEV_REASON") = Conf.DevReason
            .Value("ACT_TYPE") = Conf.ActType
            .Value("FIN_CONF") = Conf.FinalConf
            .Value("COMPLETE") = Conf.Complete
            .Value("CONF_NO") = Conf.Conf_no
        End With
        
        'Populate table for updating the serial number location table
        Set tmpSerNoLocnTable = SAPMakePPBooking.Tables.Item("SERNOLOCN")
        'Clear the temp table
        tmpSerNoLocnTable.FreeTable
        'Copy the data from the public table to the temp table
        If IntPPSerNoLocnTable.RowCount <> 0 Then
            tmpSerNoLocnTable.data = IntPPSerNoLocnTable.data
        End If
        
  

        If SAPMakePPBooking.Call = False Then
            Screen.MousePointer = vbDefault
            MsgBox SAPMakePPBooking.Exception, vbOKOnly, cnDialogTitleWorkBook
            MakeConfirmation = False
        Else
            Screen.MousePointer = vbDefault
            BookingOK = SAPMakePPBooking.Imports("BOOKING_OK")
            If BookingOK = cnYes Then 'Booking OK
                MakeConfirmation = True
            ElseIf BookingOK = cnNo Then 'Booking failed
                SAPErrorMessage = SAPMakePPBooking.Imports("ERROR_MESSAGE")
                MsgBox "BOOKING FAILED - " & SAPErrorMessage, vbExclamation, cnDialogTitleWorkBook
                MakeConfirmation = False
            Else 'Booking OK but update of serial numbers failed
                MsgBox "WARNING - UNABLE TO UPDATE THE SERIAL NUMBER LOCATIONS", vbExclamation, cnDialogTitleWorkBook
                MakeConfirmation = True
            End If
        
        End If
    
    Else 'Non PP Order
        'Set the parameters for the RFC for Non PP Orders
        With SAPMakeBooking
            .Exports("APPLICATION") = Order.Category
            .Exports("I_AUFPL") = Order.Aufpl
            .Exports("I_APLZL") = Op.Aplzl
            .Exports("SAP_DEBUG") = gblInstallation.SAPDebug
            .Exports("BACKFLUSH") = Conf.BackFlush
        End With
       

        'Set export parameters for SAPMakeBooking
   
        With SAPMakeBooking.Exports("CONF_DETAILS")
            .Value("ORDER") = Order.Number
            .Value("ACTIVITY") = Op.Number
            .Value("PERS_NO") = gblUser.ClockNumber
            .Value("WORK_CNTR") = Conf.WorkCentre
            .Value("PLANT") = Conf.Plant
            .Value("POSTG_DATE") = Conf.PostDate
            .Value("END_DATE") = Conf.EndDate
            .Value("END_TIME") = Conf.EndTime
            .Value("ACT_WORK") = Conf.ActWork
            .Value("UN_WORK") = Conf.WorkUnits
            .Value("CONF_TEXT") = Conf.ConfText
            .Value("DEV_REASON") = Conf.DevReason
            .Value("ACT_TYPE") = Conf.ActType
            .Value("FIN_CONF") = Conf.FinalConf
            .Value("COMPLETE") = Conf.Complete
        End With
  

        If SAPMakeBooking.Call = False Then
            Screen.MousePointer = vbDefault
            MsgBox SAPMakeBooking.Exception, vbOKOnly, cnDialogTitleWorkBook
            MakeConfirmation = False
        Else
            Screen.MousePointer = vbDefault
            BookingOK = SAPMakeBooking.Imports("BOOKING_OK")
            If BookingOK = cnYes Then
                MakeConfirmation = True
            Else
                SAPErrorMessage = SAPMakeBooking.Imports("ERROR_MESSAGE")
                MsgBox "BOOKING FAILED - " & SAPErrorMessage, vbExclamation, cnDialogTitleWorkBook
                MakeConfirmation = False
            End If
        
        End If
    End If
  
End Function

Public Function OpisOK2Confirm(OrdTypeCfg As OrderTypeCfgInfo, _
                               Op As OpInfo) As Boolean

   'Initialise return value to true
   OpisOK2Confirm = True

   'Check Operation Status for CNF if Order Type config demands this check
   If OrdTypeCfg.CanConfAfterFinal <> cnSAPTrue Then
      If Check4Status(OpStatusTable, cnCNF) = True Then
         OpisOK2Confirm = False
         MsgBox "OP HAS BEEN FINALLY CONFIRMED - CANNOT MAKE FURTHER BOOKINGS", _
            vbExclamation
      End If
   End If

   'Check Order Status for TECO if Order Type config demands this check
   If OrdTypeCfg.CanConfAfterTECO <> cnSAPTrue Then
      If Check4Status(OrderStatusTable, cnTECO) = True Then
         OpisOK2Confirm = False
         MsgBox "ORDER HAS BEEN TECO'd - CANNOT MAKE FURTHER BOOKINGS", _
            vbExclamation
      End If
   End If

   'Check the Op has a confirmable control key
   'If gblOrder.Category <> "10" Then ' --TCR7117--
   If Op.Bookable <> cnSAPTrue Then
         OpisOK2Confirm = False
         MsgBox "OP HAS A NON-CONFIRMABLE CONTROL KEY - CANNOT MAKE A BOOKING", _
            vbExclamation
   End If
   'End If

   End Function



Public Sub TestCICOStatus()

    'Test CICO Status as appropriate
    If gblPlantCfg.CheckClockIn = cnSAPTrue Then
        If gblUser.CICOStatus <> "I" Then
            MsgBox "WARNING - YOU ARE NOT CLOCKED IN", vbExclamation
        End If
    End If

End Sub

Public Function Check4Status(StateTable As SAPTableFactoryCtrl.Table, _
                             Status As String) As Boolean
    Dim Row As Integer
    
    'Initialise return value as False
    Check4Status = False

    'Loop thru Plant Config Table until a Plant match is found
    'then read the row into the appropriate values of UDT
    For Row = 1 To StateTable.RowCount
        If StateTable.Value(Row, "STAT") = Status And _
            StateTable.Value(Row, "INACT") <> cnSAPTrue Then
            Check4Status = True
            Exit For
        End If
    Next Row
                             

End Function

Public Function OK2DisplayForm(Reset As Boolean, _
                               TimeOut As Integer) As Boolean
                               
    Dim ViewPeriod As Single
    Static StartTime As Date
    Dim CheckTime As Date
                               
    On Error GoTo ErrorHandler
    
    'Nb the Viewperiod will be calculeted as a decimal part of a Day
    'from the TimeOut Period which is passed as a value in seconds
    
    ViewPeriod = TimeOut / 60 / 60 / 24
    
    'Test for Reset Flag
    If Reset Then
        'Set the Start Time to Now
        StartTime = Now
        OK2DisplayForm = True
    Else
        CheckTime = StartTime + ViewPeriod
        'Check that ViewTime has not Expired
        If Now > CheckTime Then
            OK2DisplayForm = False
        Else
            OK2DisplayForm = True
        End If
    
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error:" & Err.Description & " in " & Err.Source
    Exit Function
    
End Function
Public Function StripMagStripeInfo(MagStripeInfo As String, Position As String) As String

  Dim IDStartPos As Integer
  Dim IDLength As Integer
  
  ' Determine start position and length of String to retrieve from input
  IDStartPos = Mid(Position, 1, (InStr(1, Position, ",") - 1))
  IDLength = Mid(Position, (InStr(1, Position, ",") + 1))
  ' Strip out the String from the card
  StripMagStripeInfo = Mid(MagStripeInfo, IDStartPos, IDLength)
  
End Function

Public Function ConfNoWasOK() As Boolean

    If Not gblOperationWasFound Then
        ConfNoWasOK = False
        MsgBox "PLEASE ENTER THE CONFIRMATION NUMBER", vbExclamation, cnDialogTitleLogon
        frmLogon.ConfirmationNumber.SetFocus
    Else
        ConfNoWasOK = True
    End If

End Function

Public Function LoadPlantBldgs(Plant As String) As PlantBldgCfgInfo()

    Dim tmpArray() As PlantBldgCfgInfo
    Dim Row, i As Integer
    
    For Row = 1 To PlantBldgCfgTable.RowCount
        If PlantBldgCfgTable.Value(Row, "PLANT") = Plant Then
            tmpArray(i).Plant = Plant
            tmpArray(i).Building = PlantBldgCfgTable.Value(Row, "BUILDING")
            i = i + 1
        End If
    Next Row
    
    LoadPlantBldgs = tmpArray
    

End Function



Public Function ReadMethod(MethodName As String, MethIndx As Integer) As Boolean

   Dim Row As Integer

   'Initialise the return value
   ReadMethod = False
   MethIndx = 0

   For Row = 1 To UBound(Method)

      If Method(Row).Name = MethodName Then
         ReadMethod = True
         MethIndx = Row
         Exit Function
      End If
      
   Next Row

End Function

Public Sub SetMethods()

   Dim MethRow As Integer
   Dim ParmRow As Integer
   Dim OldMethRow As Integer
   Dim StartRow As Integer
   Dim EndRow As Integer
   Dim Meth As MethodInfo
   Dim Parm As MethodParmInfo
   Dim Column As Object
   Dim Structure As Object
   Dim StrucName As String
   Dim ParmValidated As Boolean
   Dim i As Integer

   ReDim MethodObjs(1 To MethodsTable.RowCount, 1) As Variant
   ReDim Method(1 To MethodsTable.RowCount) As MethodInfo

   'Populating the Method array and the Method Objects array
   'nb. that the method object (SAP Func module) is at posn 0 of
   'the second element and the SAP Table object (if present)
   'is at posn 1 of the second element. The method and method object
   'arrays are constructed so that their first element indices
   'correspond for easy access

   'Initialise Local Vars
   StartRow = 1
   EndRow = MethodParmsTable.RowCount

   For MethRow = 1 To MethodsTable.RowCount

      With Meth
         .Name = MethodsTable.Value(MethRow, "NAME")
         .FuncMod = MethodsTable.Value(MethRow, "FUNCMOD")
         .TabName = MethodsTable.Value(MethRow, "TAB_NAME")
         .InStruc = MethodsTable.Value(MethRow, "IN_STRUC")
         .OutStruc = MethodsTable.Value(MethRow, "OUT_STRUC")
         .IntExt = MethodsTable.Value(MethRow, "INT_EXT")
         .SingleExe = MethodsTable.Value(MethRow, "SINGLE_EXEC")
         .MaxParms = 0
         .Executed = False
         .ImpParmCnt = 0
         .ExpParmCnt = 0
         .StartParmRow = 0
      End With

      'Process methods depending on Internal/External Indicator
      Select Case Meth.IntExt
         Case cnExternal
            'Create the Function and Table objects for each method
            Set MethodObjs(MethRow, 0) = oSAPFunctions.Add(Meth.FuncMod)
            If Meth.TabName > vbNullString Then
               Set MethodObjs(MethRow, 1) = MethodObjs(MethRow, 0).Tables.Item(Meth.TabName)
            End If

         Case cnInternal
            'Do nothing at present as internal function will be called later.
            'The methods array will be used to validate any internal functions
            'referenced by the Excel cell name
            'nb. that only one internal function is envisaged at present
            'which will return the appropriate global variables to the Excel cell
            'the global variables will have already been established
      End Select

      'Set the Parameter values for the Method
      For ParmRow = StartRow To EndRow

         Call ReadMethParmVals(Parm, ParmRow)
         'debug.print Meth.Name, Parm.Name

         If Meth.Name = Parm.MethodName Then
         
            'initialise the test boolean
            ParmValidated = False
            
            'Record the first parameter row for this method
            'nb. that the parameter table is sorted by the Method Name
            If MethRow <> OldMethRow Then
               OldMethRow = MethRow
               Meth.StartParmRow = ParmRow
            End If

            'Update the parameter counters for this method
            Select Case Parm.ImpExp
               Case cnExport
                  Meth.ExpParmCnt = Meth.ExpParmCnt + 1
               Case cnImport
                  Meth.ImpParmCnt = Meth.ImpParmCnt + 1
            End Select

            'Update the max parameter posn as necessary
            If Meth.MaxParms < Val(Parm.Posn) Then
               Meth.MaxParms = Val(Parm.Posn)
            End If

            'Validate that the parameter from the definition table
            'actually exists in the appropriate structure, but
            'only for external methods
            Select Case Meth.IntExt

               Case cnInternal
                  'default the parmvalidated to true
                  ParmValidated = True

               Case cnExternal
                  Select Case Parm.StrTab
                     Case cnStructure
                        Select Case Parm.ImpExp
                           Case cnImport
                              StrucName = Meth.InStruc
                           Case cnExport
                              StrucName = Meth.OutStruc
                        End Select
                        'check that the structure name is not null
                        If StrucName = vbNullString Then
                           'error between method and parameter definitions
                        Else
                           'loop thru the structure fields and check that the parm name exists
                           Set Structure = MethodObjs(MethRow, 0).Imports.Item(StrucName)
                           For i = 1 To Structure.ColumnCount
                              If Structure.ColumnName(i) = Parm.Name Then
                                 ParmValidated = True
                                 Exit For
                              End If
                           Next i
                        End If

                     Case cnTable
                        If Meth.TabName = vbNullString Then
                           'Error parameter for table when method does not have a table
                        Else
                           'loop thru the columns of the table and check that
                           'the parameter exists as a column name
                           For i = 1 To MethodObjs(MethRow, 1).ColumnCount
                              Set Column = MethodObjs(MethRow, 1).Columns.Item(i)
                              If Column.Name = Parm.Name Then
                                 ParmValidated = True
                                 Exit For
                              End If
                           Next i
                        End If
                  End Select
            End Select

            If Not ParmValidated Then
               'Error as invalid parameter definition does not correspond
               'with the SAP function structures -
               'generate error message and delete the Method
               MsgBox "INVALID METHOD PARAMETER DEFINITION" & _
                  Chr(13) & "FOR METHOD " & Meth.Name & _
                  Chr(13) & "AND PARAMETER " & Parm.Name, vbExclamation
               Exit For
            End If

         Else
            'parameters for the next method so exit the loop
            'and reset the starting row for the next method
            StartRow = ParmRow
            Exit For

         End If

      Next ParmRow

      'Update the method array ONLY if the parameters validated OK
      If ParmValidated Then
         With Method(MethRow)
            .Executed = Meth.Executed
            .ExpParmCnt = Meth.ExpParmCnt
            .FuncMod = Meth.FuncMod
            .ImpParmCnt = Meth.ImpParmCnt
            .InStruc = Meth.InStruc
            .IntExt = Meth.IntExt
            .MaxParms = Meth.MaxParms
            .Name = Meth.Name
            .OutStruc = Meth.OutStruc
            .SingleExe = Meth.SingleExe
            .StartParmRow = Meth.StartParmRow
            .TabName = Meth.TabName
         End With
      End If

   Next MethRow

   'tidy up on exit
   Set Column = Nothing
   Set Structure = Nothing

End Sub

Public Function ExecMethod(MethIndx As Integer) As Boolean


   Dim Row As Integer
   Dim StartRow As Integer
   Dim EndRow As Integer
   Dim Cell As Integer
   Dim PrevCell As Integer
   Dim Parm As MethodParmInfo
   Dim Meth As MethodInfo
   Dim NewRow As Object

   'Initialise Local Vars
   PrevCell = 0

   'Read the Meth attributes and set the parm start and end rows
   Call ReadMethVals(Meth, MethIndx)
   'debug.print Meth.Name
   StartRow = Meth.StartParmRow
   EndRow = Meth.StartParmRow + Meth.ExpParmCnt + Meth.ImpParmCnt - 1

   Select Case Meth.IntExt

      Case cnExternal

         'Populate the input structure (hard-coded substitution of
         'existing globals values in SFDC)
         'use this Method until macro substitution problem is resolved
         'nb. that all methods MUST have the following fields
         'in the input structure

         With MethodObjs(MethIndx, 0).Exports(Meth.InStruc)
            .Value("SALES_ORD") = gblOrder.SalesOrder
            .Value("ROOT_FLOCN") = gblOrder.RootFLocn
            .Value("ZENGTY") = gblOrder.EngineType
            .Value("ZENGMK") = gblOrder.EngMark
            .Value("ZENGVR") = gblOrder.EngVar
            .Value("WORK_CNTR") = gblOperation.WorkCentre
            .Value("PLANT") = gblOrder.Plant
            .Value("CONFNO") = gblOperation.ConfNo
            .Value("ROOT_NETW") = gblOrder.RootNetWrk
            .Value("PM_ORDNO") = gblOrder.Number
            .Value("SO_EQUIP") = gblOrder.SOEqipment
         End With

         'Now populate the Method Table but only if a table name
         'is defined for the Method else execute the Method immediately
         If Not Meth.TabName > vbNullString Then
            GoTo ExecuteNow
         End If

         'Clear the Methods Table
         MethodObjs(MethIndx, 1).FreeTable

         'Loop thru the Named Cells array and process only those cells
         'for the Meth being executed
         For Cell = 1 To UBound(NamedCells, 1)
            'Check for Meth being processed in this cell
            If NamedCells(Cell, 3) = Meth.Name Then
               'Loop thru the Meth Parameters and build the imports to the function
               For Row = StartRow To EndRow
                  'Read the parameter table values based on the parm table row
                  Call ReadMethParmVals(Parm, Row)
                  If Parm.ImpExp = cnImport Then
                     Select Case Parm.StrTab
                        Case cnStructure
                           If Not Parm.NotImpt = cnSAPTrue Then
                              MethodObjs(MethIndx, 0).Exports(Meth.InStruc).Value(Parm.Name) = _
                                 NamedCells(Cell, Val(Parm.Posn) + cnPosnOffset)
                           End If
                        Case cnTable
                           'test that this parameter should be imported
                           'note that import parms not imported are usually used
                           'to facilitate the table read after the function has been executed
                           If Not Parm.NotImpt = cnSAPTrue Then
                              'Build the Internal table for passing to SAP
                              'Consists of inserting rows to the table
                              'and then updating the field values from the Excel
                              'add a row to the table but only once per cell
                              If Cell <> PrevCell Then
                                 'insert a row at posn 1
                                 Set NewRow = MethodObjs(MethIndx, 1).Rows.Add
                                 'update the prevcell so following import parms don't create a new row
                                 PrevCell = Cell
                              End If
                              'update the row values from the cell parameter
                              NewRow.Value(Parm.Name) = NamedCells(Cell, Val(Parm.Posn) + cnPosnOffset)
                           End If
                     End Select
                  End If 'check for Import parmater
               Next Row ' loop thru parameters
            End If 'check for correct Meth name
         Next Cell

ExecuteNow:

         'Test SAP Connection
         If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, True) Then
            Exit Function
         End If

         'set any additional import parameters
         MethodObjs(MethIndx, 0).Exports("SAP_DEBUG") = gblInstallation.SAPDebug

         'The input structure is loaded the table is filled so....
         'Call the SAP function
         If MethodObjs(MethIndx, 0).Call = False Then
            MsgBox MethodObjs(MethIndx, 0).Exception, vbExclamation, cnDialogTitleCheckPass
            ExecMethod = False
         Else
            Method(MethIndx).Executed = True
            ExecMethod = True
         End If

      Case cnInternal
         'Internal method so....
         'Load the global variable collection
         Call LoadSAPGlobals
         Method(MethIndx).Executed = True
         ExecMethod = True

   End Select

End Function

Public Sub ReadMethParmVals(Parm As MethodParmInfo, Row As Integer)

   With Parm
      .Name = MethodParmsTable.Value(Row, "NAME")
      .DefVal = MethodParmsTable.Value(Row, "DEFVAL")
      .Excel = MethodParmsTable.Value(Row, "EXCEL")
      .ImpExp = MethodParmsTable.Value(Row, "IMP_EXP")
      .Mandatory = MethodParmsTable.Value(Row, "MANDATORY")
      .NotImpt = MethodParmsTable.Value(Row, "NO_IMP")
      .MethodName = MethodParmsTable.Value(Row, "METHOD")
      .Posn = MethodParmsTable.Value(Row, "POSN")
      .StrTab = MethodParmsTable.Value(Row, "STRTAB")
      .VBName = MethodParmsTable.Value(Row, "VB_NAME")
   End With

End Sub

Public Function ValidateParms(MethIndx As Integer, _
   CellNameElems As Variant, _
   ErrorMessage As String) As Boolean

   Dim Row As Integer
   Dim StartRow As Integer
   Dim EndRow As Integer
   Dim Method As MethodInfo
   Dim Parm As MethodParmInfo
   Dim ExportsOK As Boolean
   Dim ImportsOK As Boolean

   'Initialise Local Vars
   ValidateParms = False
   Call ReadMethVals(Method, MethIndx)
   StartRow = Method.StartParmRow
   EndRow = Method.StartParmRow + Method.ExpParmCnt + Method.ImpParmCnt - 1

   'set initial values for both test booleans
   'debug.print Method.Name
   If Method.ImpParmCnt = 0 Then
      ImportsOK = True
   Else
      ImportsOK = False
   End If

   If Method.ExpParmCnt = 0 Then
      ExportsOK = True
   Else
      ExportsOK = False
   End If

   For Row = StartRow To EndRow

      'Read the parameter values for this row
      Call ReadMethParmVals(Parm, Row)
      'debug.print Parm.Name

      Select Case Parm.ImpExp

         Case cnExport
            'Validate Export parameters by validating the name in the cell
            'against the parm name from SAP
            'nb. there should only be one export parameter at posn 1
            If UCase(CellNameElems(Val(Parm.Posn))) = Parm.Name Then
               'Export Parameter is valid
               ExportsOK = True
            End If

         Case cnImport
            'Validate Import parameters
            If Parm.Mandatory = cnSAPTrue Then
               'check that the manadatory parameter has been entered in the cell
               If CellNameElems(Val(Parm.Posn)) = vbNullString Then
                  'error message
                  ValidateParms = False
                  ErrorMessage = "MANDATORY PARAMETER @ POSN " & Parm.Posn & _
                     " IS MISSING"
                  Exit Function
               Else
                  ImportsOK = True
               End If

            Else 'optional parameter
               'populate the optional value into the cell
               'but only if there is not already a value in the cell
               If CellNameElems(Val(Parm.Posn)) = vbNullString Then
                  'debug.print UBound(CellNameElems)
                  CellNameElems(Val(Parm.Posn)) = Parm.DefVal
               End If
            End If

      End Select

   Next Row

   If Not ExportsOK Then
      ErrorMessage = "HAS AN INVALID EXPORT PARAMETER"
   End If

   'Set the function return value
   If ImportsOK And ExportsOK Then
      ValidateParms = True
   End If

End Function

Public Sub LoadNamedCellArray(wsSheet As Worksheet, DocNo As Integer)

   Dim CellName As String
   Dim CellNameElems As Variant
   Dim MethParm As MethodParmInfo
   Dim MethIndx As Integer
   Dim ElemPosn As Integer
   Dim Cell As Integer
   Dim NamedCellCount As Integer
   Dim OvMaxParmCnt As Integer
   Dim ElemCnt As Integer
   Dim MethCellCnt As Integer
   Dim MethodName As String
   Dim CellRef As String
   Dim RefCellName As String
   Dim i As Integer, j As Integer
   Dim Method As MethodInfo
   Dim tmpNamedCells() As String
   Dim CellRefFound As Boolean
   Dim ErrorMessage As String

   'Determine No. of named cells in document
   NamedCellCount = gblDocuments(DocNo).Doc.Names.Count

   'Re-Dimension the temp Named Cell array based on the named cell count
   'and allow for up to 16 parameters !!!!
   ReDim tmpNamedCells(NamedCellCount, 20)

   For Cell = 1 To NamedCellCount

      'Copy cell name to local var
      CellName = gblDocuments(DocNo).Doc.Names(Cell).Name
      'debug.print CellName

      'Check the Cell Name begins with "SAP", if not, skip to next cell
      If Not UCase(Left(CellName, 3)) = cnSAPMethod Then
         GoTo SkipToNextCell
      End If

      'Split the cell name string and update Max no of elements as necessary
      CellNameElems = Split(CellName, ".", -1)

      'Set the Method Name (converted to upper case)
      MethodName = UCase(CellNameElems(0))
      CellNameElems(0) = MethodName

      'Validate the Method Name from the SAP list
      If Not ReadMethod(MethodName, MethIndx) Then
         'this is not a valid method
         MsgBox "CELLNAME " & CellName & Chr(13) & _
            "HAS AN INVALID METHOD", vbExclamation
         'So ignore this cell and skip to the next
         GoTo SkipToNextCell
      End If

      'Read the Method values
      Call ReadMethVals(Method, MethIndx)
      'debug.print Method.Name

      'Check that the max no of parms for this Method is not exceeded
      If UBound(CellNameElems) > Method.MaxParms Then
         MsgBox "CELLNAME " & CellName & Chr(13) & _
            "HAS TOO MANY PARAMETERS", vbExclamation
         'So ignore this cell and skip to the next
         GoTo SkipToNextCell
      End If

      'Redimension the cell name elements array based on the max
      'number of parameters for this method
      ReDim Preserve CellNameElems(Method.MaxParms)

      'Set the max number of parameters as necessary
      'nb. this is used to dimension the cell array later
      If Method.MaxParms > OvMaxParmCnt Then
         OvMaxParmCnt = Method.MaxParms
      End If

      'Validate the parameters in the cell
      If Not ValidateParms(MethIndx, CellNameElems, ErrorMessage) Then
         'Ignore this cell and skip to the next
         MsgBox "CELLNAME " & CellName & Chr(13) & _
            ErrorMessage, vbExclamation
         GoTo SkipToNextCell
      End If

      'Perform substitution of cell refs where necessary
      For ElemPosn = 1 To UBound(CellNameElems)
         'debug.print Method.Name, CellNameElems(ElemPosn)
         'Test for ? at first char of the element and if so
         'navigate to and read the referenced cell for the value
         If Left(CellNameElems(ElemPosn), 1) = cnRefCellIndicator Then
            'read the remaining string, which should be a cell name
            CellRef = UCase(Mid(CellNameElems(ElemPosn), 2))
            'initialise test boolean prior checking existence of cell name
            CellRefFound = False
            'check that the cellname exists
            For i = 1 To gblDocuments(DocNo).Doc.Names.Count
               RefCellName = UCase(gblDocuments(DocNo).Doc.Names.Item(i).Name)
               If RefCellName = CellRef Then
                  CellRefFound = True
                  Exit For
               End If
            Next i
            'test whether cell name was found
            If Not CellRefFound Then
               MsgBox "CELLNAME " & CellName & Chr(13) & _
                  "CELLNAME REFERENCED DOES NOT EXIST", vbExclamation
               GoTo SkipToNextCell
            End If
            'cell name is valid so....
            'go get the value at the cell ref and substitute it into this element
            CellNameElems(ElemPosn) = gblDocuments(DocNo).Doc.Names(CellRef).RefersToRange.Value
            'check whether a value was found at the refenced cell
            If CellNameElems(ElemPosn) = vbNullString Then
               MsgBox "CELLNAME " & CellName & Chr(13) & _
                  "REFERENCED CELL IS EMPTY", vbExclamation
               GoTo SkipToNextCell
            End If
         End If
      Next ElemPosn

      'If we've reached this point it means everything validated OK
      'so update the Valid Cell count, which we'll use for dimensioning
      'the 2D array
      MethCellCnt = MethCellCnt + 1

      'Load the 2D array with the contents of the 1D array
      'nb. posn 1 of the second index will hold the returned value
      tmpNamedCells(MethCellCnt, 0) = CellName
      tmpNamedCells(MethCellCnt, 2) = MethIndx
      For i = 0 To UBound(CellNameElems)
         tmpNamedCells(MethCellCnt, i + cnPosnOffset) = UCase(CellNameElems(i))
      Next i

SkipToNextCell:
   Next Cell

   'Now we've processed all the named cells
   'Redimension the global 2D array based on
   'the no of valid methods and max number of elements found
   ReDim NamedCells(1 To MethCellCnt, OvMaxParmCnt + cnPosnOffset)

   'and copy the contents of the temp array into it
   For i = 1 To MethCellCnt
      For j = 0 To OvMaxParmCnt + cnPosnOffset
         NamedCells(i, j) = tmpNamedCells(i, j)
      Next j
   Next i

   'store the method cell count to the global var
   gblMethCellCnt = MethCellCnt

   'then erase the temp array
   Erase tmpNamedCells

End Sub

Public Sub ReadMethVals(Meth As MethodInfo, MethIndx As Integer)

   With Meth
      .ExpParmCnt = Method(MethIndx).ExpParmCnt
      .FuncMod = Method(MethIndx).FuncMod
      .ImpParmCnt = Method(MethIndx).ImpParmCnt
      .InStruc = Method(MethIndx).InStruc
      .IntExt = Method(MethIndx).IntExt
      .Name = Method(MethIndx).Name
      .OutStruc = Method(MethIndx).OutStruc
      .StartParmRow = Method(MethIndx).StartParmRow
      .TabName = Method(MethIndx).TabName
      .MaxParms = Method(MethIndx).MaxParms
      .SingleExe = Method(MethIndx).SingleExe
      .Executed = Method(MethIndx).Executed
   End With

End Sub

Public Sub ProcessNamedCells()

   'This routine actually populates the export value into the named cell array
   'a later routine will substitute the export value into the named cell
   '(See UpdateNamedCells function)

   Dim Cell As Integer
   Dim i As Integer, j As Integer
   Dim StartRow As Integer, EndRow As Integer
   Dim MethIndx As Integer
   Dim Method As MethodInfo
   Dim Parm As MethodParmInfo
   Dim ExportParm As String
   Dim ExportMedium As String * 1
   Dim ImportParms() As String
   Dim ImptParmCnt As Integer
   Dim CellParmName As String
   Dim TableRow As Object
   Dim RowFound As Boolean
   Dim FoundRow As Integer

   For Cell = 1 To UBound(NamedCells, 1)
      'debug.print NamedCells(Cell, 0)
      'set the local var for method index from posn 2
      'of the second element of the named cell array
      MethIndx = NamedCells(Cell, 2)
      'Read the Method Attributes for this cell using method index
      Call ReadMethVals(Method, MethIndx)

      'Execute Methods as necessary by .....
      'determining if the Method is a "single execution"
      If Method.SingleExe = cnSAPTrue Then
         'check if the method has already been executed
         'and if not, then execute it
         If Not Method.Executed Then
            If ExecMethod(MethIndx) Then
               Method.Executed = True
            Else
               Method.Executed = False
            End If
         End If
      Else 'this a multiple execution method so....
         'execute it regardless
         If ExecMethod(MethIndx) Then
            Method.Executed = True
         Else
            Method.Executed = False
         End If
      End If

      'now the method has been executed so...
      'we can evaluate the parameters and determine the export value

      'first determine the parameter table rows for this method
      'and loop thru them
      StartRow = Method.StartParmRow
      EndRow = Method.StartParmRow + Method.ExpParmCnt + Method.ImpParmCnt - 1
      're-dimension the temp import parm array
      'nb. redim will clear the data for each cell
      'and reset the import parm counter
      'but ONLY if the method has any import parameters
      If Method.ImpParmCnt > 0 Then
         ReDim ImportParms(1 To Method.ImpParmCnt, 1)
         ImptParmCnt = 1
      End If

      For i = StartRow To EndRow
         'read the parameter attributes
         'note that the export parameters will be read first as the
         'table was sorted by Method and Import/Export indicator
         Call ReadMethParmVals(Parm, i)
         Select Case Parm.ImpExp

            Case cnExport
               'should only find one export parameter, with a matching name,
               'for any given method
               CellParmName = NamedCells(Cell, Val(Parm.Posn) + cnPosnOffset)
               If Parm.Name = CellParmName Then
                  ExportParm = Parm.Name
                  ExportMedium = Parm.StrTab
               End If

            Case cnImport
               'Populate the temp import parameters array
               'posn 0 = parm name, posn 1 = parm value
               ImportParms(ImptParmCnt, 0) = Parm.Name
               ImportParms(ImptParmCnt, 1) = NamedCells(Cell, Val(Parm.Posn) + cnPosnOffset)
               ImptParmCnt = ImptParmCnt + 1

         End Select

      Next i

      Select Case Method.IntExt
         Case cnExternal
            'Process based on the export medium (Table or Structure)
            Select Case ExportMedium

               Case cnStructure
                  'export parameter is in the export structure so
                  'just read the export structure for this parm
                  NamedCells(Cell, 1) = MethodObjs(MethIndx, 0).Imports(Method.OutStruc).Value(ExportParm)

               Case cnTable
                  'the export parameter is in the table so...
                  'read all the rows in the table and find the one
                  'which matches the import parameters
                  RowFound = False
                  For i = 1 To MethodObjs(MethIndx, 1).RowCount
                     'create a row object
                     Set TableRow = MethodObjs(MethIndx, 1).Rows.Item(i)
                     'check for matching import values in the row
                     For j = 1 To Method.ImpParmCnt
                        If Not TableRow.Value(ImportParms(j, 0)) = ImportParms(j, 1) Then
                           RowFound = False
                           Exit For
                        Else
                           RowFound = True
                        End If
                     Next j
                     'if the row was found then set the row pointer and
                     'exit from the table loop else continue to next row
                     If RowFound = True Then
                        FoundRow = i
                        Exit For
                     End If
                  Next i

                  'Set the export value if the row was found
                  'and set the Row Read flag in the table
                  If RowFound Then
                     NamedCells(Cell, 1) = TableRow.Value(ExportParm)
                     TableRow.Value("ROW_READ") = cnSAPTrue
                  End If

            End Select

         Case cnInternal
            'execute the Internal Method to populate the export value
            NamedCells(Cell, 1) = SAPGlobalVars.Item(ExportParm)

      End Select

   Next Cell
   
   Call CheckAllRowsRead


   'tidy up on exit
   Set TableRow = Nothing

End Sub

Public Sub UpdateNamedCells(wsSheet As Worksheet, DocNo As Integer)

   Dim Cell As Integer

   '=======================================================
   ' Unprotect sheet to allow update of named cells
   '=======================================================
   Call gblDocuments(DocNo).Doc.ActiveSheet.Unprotect

   For Cell = 1 To UBound(NamedCells, 1)
      gblDocuments(DocNo).Doc.Names(NamedCells(Cell, 0)).RefersToRange.Value = NamedCells(Cell, 1)
   Next Cell

   '==================
   ' Protect the sheet
   '==================
   Call gblDocuments(DocNo).Doc.ActiveSheet.Protect("", True, True, True)

End Sub

Public Sub LoadSAPGlobals()

   Set SAPGlobalVars = New Collection

   SAPGlobalVars.Add gblOrder.Customer, "CUSTOMER"
   SAPGlobalVars.Add gblOrder.SalesOrder, "SALES_ORD"
   SAPGlobalVars.Add gblOrder.EngSerialNo, "ENG_SERNO"
   SAPGlobalVars.Add gblOrder.ModuleNumber, "MOD_SERNO"
   SAPGlobalVars.Add gblOrder.EngineType, "ENG_TYPE"
   SAPGlobalVars.Add gblOrder.EngMark, "ENG_MARK"
   SAPGlobalVars.Add gblOrder.EngVar, "ENG_VAR"
   SAPGlobalVars.Add gblOrder.RootFLocn, "ROOT_FLOCN"
   SAPGlobalVars.Add gblOrder.Operator, "OPERATOR"
   SAPGlobalVars.Add gblOrder.ICAOCode, "ICAO_CODE"
   SAPGlobalVars.Add gblOrder.DISLevel, "DIS_LEV"
   SAPGlobalVars.Add gblOrder.AircraftType, "ACTYP"
   SAPGlobalVars.Add gblOrder.Authority, "AUTHY"
   SAPGlobalVars.Add gblOrder.ExAircraftTailNo, "EXATN"

End Sub

Public Sub CheckAllRowsRead()

   Dim i As Integer, j As Integer, k As Integer
   Dim StartRow As Integer, EndRow As Integer
   Dim Meth As MethodInfo, Parm As MethodParmInfo
   Dim WarningMess As String
   Dim ParmValue As Variant

   'Check that all rows in all tables have been read
   For i = 1 To MethodsTable.RowCount
      Call ReadMethVals(Meth, i)
      If Meth.Executed Then
         If Meth.TabName > vbNullString Then
            'set the method parameters rows for import values
            StartRow = Meth.StartParmRow + Meth.ExpParmCnt
            EndRow = StartRow + Meth.ImpParmCnt
            For j = 1 To MethodObjs(i, 1).RowCount
            'Debug.Print
               If MethodObjs(i, 1).Value(j, "ROW_READ") <> cnSAPTrue Then
                  'warning this row has not been read
                  'output the method name and import paramater names and values
                  WarningMess = "method=" & Meth.Name
                  For k = StartRow To EndRow
                     Call ReadMethParmVals(Parm, k)
                     If Parm.StrTab = cnTable Then
                        ParmValue = MethodObjs(i, 1).Value(j, Parm.Name)
                        WarningMess = WarningMess & "  " & LCase(Parm.Name) & "=" & ParmValue
                     End If
                  Next k
                  'output message
                  MsgBox "MORE DATA RETRIEVED THAN CAN BE DISPLAYED - TRUNCATION HAS OCCURRED" _
                      & Chr(13) & WarningMess, vbExclamation
               End If
            Next j
         End If
      End If
   Next i

End Sub

Public Sub ClearFlexGrid(ByRef FlxGrd As MSFlexGrid)

   On Error GoTo ExitOnError

   Dim i As Integer
   Dim Rows As Integer
   
   Rows = FlxGrd.Rows

   For i = 2 To Rows + 1

      FlxGrd.RemoveItem (2)

   Next i

ExitOnError:

   ' continue

End Sub

Public Function CanWorkBook() As Boolean

'Initialise the return value
CanWorkBook = True

   'Check the Person is authorised to Work Book via SFDC
   If gblUser.CanWorkBook <> cnSAPTrue Then
      MsgBox "YOU ARE NOT AUTHORISED TO WORK BOOK VIA SFDC", vbExclamation, cnDialogTitleWorkBook
      CanWorkBook = False
   End If
   
End Function


