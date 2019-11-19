Attribute VB_Name = "launcher"
Attribute VB_HelpID = 850
'Внешние функции



Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type Point
    X As Long
    Y As Long
End Type

Public Const HWND_TOPMOST = -1
Attribute HWND_TOPMOST.VB_VarHelpID = 940
Public Const SWP_NOSIZE = &H1
Attribute SWP_NOSIZE.VB_VarHelpID = 1225
Public Const HWND_TOP = 0
Attribute HWND_TOP.VB_VarHelpID = 935
Public Const SWP_NOMOVE = &H2
Attribute SWP_NOMOVE.VB_VarHelpID = 1220
Public Const SWP_SHOWWINDOW = &H40
Attribute SWP_SHOWWINDOW.VB_VarHelpID = 1235
Public Const GWL_STYLE = (-16)
Attribute GWL_STYLE.VB_VarHelpID = 900
Public Const SWP_NOZORDER = &H4
Attribute SWP_NOZORDER.VB_VarHelpID = 1230
Public Const SWP_FRAMECHANGED = &H20
Attribute SWP_FRAMECHANGED.VB_VarHelpID = 1215
   
Public Const SW_HIDE = 0
Attribute SW_HIDE.VB_VarHelpID = 1160
Public Const SW_MAXIMIZE = 3
Attribute SW_MAXIMIZE.VB_VarHelpID = 1165
Public Const SW_MINIMIZE = 6
Attribute SW_MINIMIZE.VB_VarHelpID = 1170
Public Const SW_RESTORE = 9
Attribute SW_RESTORE.VB_VarHelpID = 1175
Public Const SW_SHOW = 5
Attribute SW_SHOW.VB_VarHelpID = 1180
Public Const SW_SHOWMAXIMIZED = 3
Attribute SW_SHOWMAXIMIZED.VB_VarHelpID = 1185
Public Const SW_SHOWMINIMIZED = 2
Attribute SW_SHOWMINIMIZED.VB_VarHelpID = 1190
Public Const SW_SHOWMINNOACTIVE = 7
Attribute SW_SHOWMINNOACTIVE.VB_VarHelpID = 1195
Public Const SW_SHOWNA = 8
Attribute SW_SHOWNA.VB_VarHelpID = 1200
Public Const SW_SHOWNOACTIVATE = 4
Attribute SW_SHOWNOACTIVATE.VB_VarHelpID = 1205
Public Const SW_SHOWNORMAL = 1
Attribute SW_SHOWNORMAL.VB_VarHelpID = 1210


' Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Attribute DrawText.VB_HelpID = 855
Private Declare Function GetScrollPos Lib "User32.dll" (ByVal hwnd As Long, ByVal nBar As Integer) As Integer
Private Declare Function PostMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal WParm As Long, ByVal lParm As Integer) As Long



Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Attribute SetWindowPos.VB_HelpID = 1140
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Attribute GetUserName.VB_HelpID = 895
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Attribute FlashWindow.VB_HelpID = 875

Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Attribute GetCaretBlinkTime.VB_HelpID = 880
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Attribute GetCurrentProcess.VB_HelpID = 885
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Attribute TerminateProcess.VB_HelpID = 1240



Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, IpData As NOTIFYICONDATA) As Long
Attribute Shell_NotifyIcon.VB_HelpID = 1145

Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Attribute NIM_ADD.VB_VarHelpID = 1000
Public Const NIM_MODIFY = &H1
Attribute NIM_MODIFY.VB_VarHelpID = 1010
Public Const NIM_DELETE = &H2
Attribute NIM_DELETE.VB_VarHelpID = 1005
Public Const NIF_MESSAGE = &H1
Attribute NIF_MESSAGE.VB_VarHelpID = 990
Public Const NIF_ICON = &H2
Attribute NIF_ICON.VB_VarHelpID = 985
Public Const NIF_TIP = &H4
Attribute NIF_TIP.VB_VarHelpID = 995
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Attribute NIF_DOALL.VB_VarHelpID = 980
Public Const WM_MOUSEMOVE = &H200
Attribute WM_MOUSEMOVE.VB_VarHelpID = 1255
Public Const WM_LBUTTONDBLCLK = &H203
Attribute WM_LBUTTONDBLCLK.VB_VarHelpID = 1245
Public Const WM_LBUTTONDOWN = &H201
Attribute WM_LBUTTONDOWN.VB_VarHelpID = 1250
Public Const WM_RBUTTONDOWN = &H204
Attribute WM_RBUTTONDOWN.VB_VarHelpID = 1260

Public Const HKEY_CLASSES_ROOT = &H80000000
Attribute HKEY_CLASSES_ROOT.VB_VarHelpID = 905
Public Const HKEY_CURRENT_USER = &H80000001
Attribute HKEY_CURRENT_USER.VB_VarHelpID = 915
Public Const HKEY_LOCAL_MACHINE = &H80000002
Attribute HKEY_LOCAL_MACHINE.VB_VarHelpID = 925
Public Const HKEY_USERS = &H80000003
Attribute HKEY_USERS.VB_VarHelpID = 930
Public Const HKEY_CURRENT_CONFIG = &H80000005
Attribute HKEY_CURRENT_CONFIG.VB_VarHelpID = 910
Public Const HKEY_DYN_DATA = &H80000006
Attribute HKEY_DYN_DATA.VB_VarHelpID = 920

'Registry Specific Access Rights
Public Const KEY_QUERY_VALUE = &H1
Attribute KEY_QUERY_VALUE.VB_VarHelpID = 970
Public Const KEY_SET_VALUE = &H2
Attribute KEY_SET_VALUE.VB_VarHelpID = 975
Public Const KEY_CREATE_SUB_KEY = &H4
Attribute KEY_CREATE_SUB_KEY.VB_VarHelpID = 955
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Attribute KEY_ENUMERATE_SUB_KEYS.VB_VarHelpID = 960
Public Const KEY_NOTIFY = &H10
Attribute KEY_NOTIFY.VB_VarHelpID = 965
Public Const KEY_CREATE_LINK = &H20
Attribute KEY_CREATE_LINK.VB_VarHelpID = 950
Public Const KEY_ALL_ACCESS = &H3F
Attribute KEY_ALL_ACCESS.VB_VarHelpID = 945

'Open/Create Options
Public Const REG_OPTION_NON_VOLATILE = 0&
Attribute REG_OPTION_NON_VOLATILE.VB_VarHelpID = 1070
Public Const REG_OPTION_VOLATILE = &H1
Attribute REG_OPTION_VOLATILE.VB_VarHelpID = 1075

'Key creation/open disposition
Public Const REG_CREATED_NEW_KEY = &H1
Attribute REG_CREATED_NEW_KEY.VB_VarHelpID = 1020
Public Const REG_OPENED_EXISTING_KEY = &H2
Attribute REG_OPENED_EXISTING_KEY.VB_VarHelpID = 1065

'masks for the predefined standard access types
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Attribute STANDARD_RIGHTS_ALL.VB_VarHelpID = 1155
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Attribute SPECIFIC_RIGHTS_ALL.VB_VarHelpID = 1150

'Define severity codes
Public Const ERROR_SUCCESS = 0&
Attribute ERROR_SUCCESS.VB_VarHelpID = 870
Public Const ERROR_ACCESS_DENIED = 5
Attribute ERROR_ACCESS_DENIED.VB_VarHelpID = 860
Public Const ERROR_NO_MORE_ITEMS = 259
Attribute ERROR_NO_MORE_ITEMS.VB_VarHelpID = 865

'Predefined Value Types

'No value type
Public Const REG_NONE = (0)
Attribute REG_NONE.VB_VarHelpID = 1060
'Unicode nul terminated string
Public Const REG_SZ = (1)
Attribute REG_SZ.VB_VarHelpID = 1090
'Unicode nul terminated string w/enviornment var
Public Const REG_EXPAND_SZ = (2)
Attribute REG_EXPAND_SZ.VB_VarHelpID = 1040
'Free form binary
Public Const REG_BINARY = (3)
Attribute REG_BINARY.VB_VarHelpID = 1015
'32-bit number
Public Const REG_DWORD = (4)
Attribute REG_DWORD.VB_VarHelpID = 1025
'32-bit number (same as REG_DWORD)
Public Const REG_DWORD_LITTLE_ENDIAN = (4)
Attribute REG_DWORD_LITTLE_ENDIAN.VB_VarHelpID = 1035
'32-bit number
Public Const REG_DWORD_BIG_ENDIAN = (5)
Attribute REG_DWORD_BIG_ENDIAN.VB_VarHelpID = 1030
'Symbolic Link (unicode)
Public Const REG_LINK = (6)
Attribute REG_LINK.VB_VarHelpID = 1050
'Multiple Unicode strings
Public Const REG_MULTI_SZ = (7)
Attribute REG_MULTI_SZ.VB_VarHelpID = 1055
'Resource list in the resource map
Public Const REG_RESOURCE_LIST = (8)
Attribute REG_RESOURCE_LIST.VB_VarHelpID = 1080
'Resource list in the hardware description
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)
Attribute REG_FULL_RESOURCE_DESCRIPTOR.VB_VarHelpID = 1045
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)
Attribute REG_RESOURCE_REQUIREMENTS_LIST.VB_VarHelpID = 1085


'Structures Needed For Registry Prototypes
Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'Registry Functions
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Attribute RegCloseKey.VB_HelpID = 1095
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Attribute RegCreateKeyEx.VB_HelpID = 1100
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Attribute RegDeleteKey.VB_HelpID = 1105
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Attribute RegDeleteValue.VB_HelpID = 1110
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Attribute RegEnumKeyEx.VB_HelpID = 1115
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Attribute RegEnumValue.VB_HelpID = 1120
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Attribute RegOpenKeyEx.VB_HelpID = 1125
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Attribute RegQueryValueEx.VB_HelpID = 1130
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Attribute RegSetValueEx.VB_HelpID = 1135

'Имя пользователя Windows
'Parameters:
' параметров нет
'Returns:
'  значение типа String
'See Also:
'  DrawText
'  ERROR_ACCESS_DENIED
'  ERROR_NO_MORE_ITEMS
'  ERROR_SUCCESS
'  FlashWindow
'  GetCaretBlinkTime
'  GetCurrentProcess
'  GetUserName
'  GWL_STYLE
'  HKEY_CLASSES_ROOT
'  HKEY_CURRENT_CONFIG
'  HKEY_CURRENT_USER
'  HKEY_DYN_DATA
'  HKEY_LOCAL_MACHINE
'  HKEY_USERS
'  HWND_TOP
'  HWND_TOPMOST
'  KEY_ALL_ACCESS
'  KEY_CREATE_LINK
'  KEY_CREATE_SUB_KEY
'  KEY_ENUMERATE_SUB_KEYS
'  KEY_NOTIFY
'  KEY_QUERY_VALUE
'  KEY_SET_VALUE
'  NIF_DOALL
'  NIF_ICON
'  NIF_MESSAGE
'  NIF_TIP
'  NIM_ADD
'  NIM_DELETE
'  NIM_MODIFY
'  REG_BINARY
'  REG_CREATED_NEW_KEY
'  REG_DWORD
'  REG_DWORD_BIG_ENDIAN
'  REG_DWORD_LITTLE_ENDIAN
'  REG_EXPAND_SZ
'  REG_FULL_RESOURCE_DESCRIPTOR
'  REG_LINK
'  REG_MULTI_SZ
'  REG_NONE
'  REG_OPENED_EXISTING_KEY
'  REG_OPTION_NON_VOLATILE
'  REG_OPTION_VOLATILE
'  REG_RESOURCE_LIST
'  REG_RESOURCE_REQUIREMENTS_LIST
'  REG_SZ
'  RegCloseKey
'  RegCreateKeyEx
'  RegDeleteKey
'  RegDeleteValue
'  RegEnumKeyEx
'  RegEnumValue
'  RegOpenKeyEx
'  RegQueryValueEx
'  RegSetValueEx
'  SetWindowPos
'  Shell_NotifyIcon
'  SPECIFIC_RIGHTS_ALL
'  STANDARD_RIGHTS_ALL
'  SW_HIDE
'  SW_MAXIMIZE
'  SW_MINIMIZE
'  SW_RESTORE
'  SW_SHOW
'  SW_SHOWMAXIMIZED
'  SW_SHOWMINIMIZED
'  SW_SHOWMINNOACTIVE
'  SW_SHOWNA
'  SW_SHOWNOACTIVATE
'  SW_SHOWNORMAL
'  SWP_FRAMECHANGED
'  SWP_NOMOVE
'  SWP_NOSIZE
'  SWP_NOZORDER
'  SWP_SHOWWINDOW
'  TerminateProcess
'  WM_LBUTTONDBLCLK
'  WM_LBUTTONDOWN
'  WM_MOUSEMOVE
'  WM_RBUTTONDOWN
'Example:
' dim variable as String
'  variable = me.GetUser()
Public Function GetUser() As String
Attribute GetUser.VB_HelpID = 890
  Dim sBuffer As String
  Dim lSize As Long
  Dim mUserName  As String
  sBuffer = Space$(255)
  lSize = Len(sBuffer)
  Call GetUserName(sBuffer, lSize)
  GetUser = Left$(sBuffer, lSize)
End Function
