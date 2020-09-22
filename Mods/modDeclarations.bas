Attribute VB_Name = "modDeclarations"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Running                           As Boolean
Public Type SHFILEOPSTRUCT
    hWnd                                   As Long
    wFunc                                  As Long
    pFrom                                  As String
    pTo                                    As String
    fFlags                                 As Integer
    fAnyOperationsAborted                  As Boolean
    hNameMappings                          As Long
    lpszProgressTitle                      As String
End Type
Public Const FO_DELETE                   As Long = &H3
Public Type IconType
    cbSize                                 As Long
    picType                                As PictureTypeConstants
    hIcon                                  As Long
End Type
Public Type CLSIdType
    Id(16)                                 As Byte
End Type
Public Type ShellFileInfoType
    hIcon                                  As Long
    iIcon                                  As Long
    dwAttributes                           As Long
    szDisplayName                          As String * 260
    szTypeName                             As String * 80
End Type
Public Const Large                       As Long = &H100
Public Const VER_PLATFORM_WIN32_NT       As Integer = 2
Public Type OSVERSIONINFO
    dwOSVersionInfoSize                    As Long
    dwMajorVersion                         As Long
    dwMinorVersion                         As Long
    dwBuildNumber                          As Long
    dwPlatformId                           As Long
    szCSDVersion                           As String * 128
End Type
Private Type TypeSignature
    SignatureFilename                      As String
    SignatureDate                          As String
    SignatureOnlineFilename                As String
    SignatureCount                         As Integer
End Type
Public Enum RM
    Normal = 0
    TrayOnly = 1
    ScanFile = 3
    SecureFile = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Normal, TrayOnly, ScanFile
#End If
#If False Then
Private Normal, TrayOnly, ScanFile
#End If
Private Type AntiVirus
    AVname                                 As String
    Runmode                                As RM
    Signature                              As TypeSignature
End Type
Public AV                                As AntiVirus
Private Type SHItemID
    cb                                     As Long
    abID                                   As Byte
End Type
Public Type ItemIDList
    mkid                                   As SHItemID
End Type
Public Type BROWSEINFO
    hOwner                                 As Long
    pidlRoot                               As Long
    pszDisplayName                         As String
    lpszTitle                              As String
    ulFlags                                As Long
    lpCallbackProc                         As Long
    lParam                                 As Long
    iImage                                 As Long
End Type
Public Enum VirusT
    Executable = 0
    Script = 1
End Enum
Public Enum pStatus
    Max = 1
    Min = 0
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Executable, Script
#End If
Private Type TypeVirus
    FileNameShort                           As String
    Reason                                 As String
    FileName                               As String
Type                                   As VirusT
End Type
Public Virus                             As TypeVirus
#If Win16 Then
Public Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, _
                                            ByVal hWndInsertAfter As Integer, _
                                            ByVal X As Integer, _
                                            ByVal Y As Integer, _
                                            ByVal cx As Integer, _
                                            ByVal cy As Integer, _
                                            ByVal wFlags As Integer)
#Else
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long
#End If
Public Type FILETIME
    dwLowDateTime                          As Long
    dwHighDateTime                         As Long
End Type
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Type WIN32_FIND_DATA
    dwFileAttributes                       As Long
    ftCreationTime                         As FILETIME
    ftLastAccessTime                       As FILETIME
    ftLastWriteTime                        As FILETIME
    nFileSizeHigh                          As Long
    nFileSizeLow                           As Long
    dwReserved0                            As Long
    dwReserved1                            As Long
    cFileName                              As String * 260
    cAlternate                             As String * 14
End Type
Public Const KEY_ALL_ACCESS = &H3F
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const REG_PRIMARY_KEY = "Software\Classes\"
Public Const REG_SHELL_KEY = "Shell\"
Public Const REG_SHELL_OPEN_KEY = "Open\"
Public Const REG_SHELL_OPEN_COMMAND_KEY = "Command"
Public Const REG_ICON_KEY = "DefaultIcon"
Public Const REG_SZ = 1
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const ERROR_SUCCESS = 0&
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As Any, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Public SH                                As New Shell    'reference to shell32.dll class
Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, _
                                                                     riid As CLSIdType, _
                                                                     ByVal fown As Long, _
                                                                     lpUnk As Object) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                ByVal dwFileAttributes As Long, _
                                                                                psfi As ShellFileInfoType, _
                                                                                ByVal cbFileInfo As Long, _
                                                                                ByVal uFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetFileNameFromBrowseW Lib "Shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As Long, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As Long, _
                                                                          ByVal lpstrDefExt As Long, _
                                                                          ByVal lpstrFilter As Long, _
                                                                          ByVal lpstrTitle As Long) As Long
Public Declare Function GetFileNameFromBrowseA Lib "Shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As String, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As String, _
                                                                          ByVal lpstrDefExt As String, _
                                                                          ByVal lpstrFilter As String, _
                                                                          ByVal lpstrTitle As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub FindClose Lib "kernel32" (ByVal hFindFile As Long)
Public Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, _
                                                       lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, _
                                                      lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long


