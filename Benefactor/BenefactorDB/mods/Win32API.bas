Attribute VB_Name = "modWin32API"
'-------------------------------------------------------------------------------
'Module Name: modWin32API
'Author: Dave Richkun
'Date: 11/08/1999
'Description: This module contains definitions and supporting code for Win32
'             API declarations.  The Win32 calls may be called from any
'             application.
'-------------------------------------------------------------------------------
'Revision History:
'
'-------------------------------------------------------------------------------
Option Explicit

'--------------------------
'Windows Constants
'-------------------------

Public Const GWL_STYLE = (-16)
Public Const WS_CHILD = &H40000000

Public Const SYNCHRONIZE = 1048576
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&

Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const NERR_BASE = 2100
Public Const MAX_NERR = NERR_BASE + 899
Public Const LOAD_LIBRARY_AS_DATAFILE = &H2

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type


'--------------------------
'Windows Declarations
'-------------------------

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long
    
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, _
    ByVal lpEnumFunc As Long, ByVal lParam&) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
     ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
    
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, _
    ByVal idAttachTo As Long, ByVal fAttach As Long) As Long

Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    lpEnvironment As Any, ByVal lpCurrentDriectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function OpenProcess Lib "kernel32.dll" _
    (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" _
    (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long

Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Declare Function NetApiBufferFree& Lib "netapi32" (ByVal Buffer As Long)

Public Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)

Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, _
     ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
     Arguments As Any) As Long

Public Declare Function NetUserSetInfo Lib "netapi32.dll" (ByVal ServerName As String, _
    ByVal Username As String, ByVal Level As Long, UserInfo As Any, ParmError As Long) As Long

Public Declare Function NetGetDCName Lib "netapi32.dll" (ServerName As Long, _
    DomainName As Byte, bufptr As Long) As Long

Public Declare Function NetUserChangePassword Lib "netapi32.dll" (ByVal DomainName As String, _
    ByVal Username As String, ByVal OldPassword As String, ByVal NewPassword As String) As Long


Public Function GetLoginName() As String

    Dim strBuffer As String * 200
    Dim lngLength As Long
    Dim lngReturnCode As Long

    lngLength = 199
    lngReturnCode = GetUserName(strBuffer, lngLength)
    
    GetLoginName = Left$(strBuffer, lngLength - 1)
    
End Function


Public Function GetPDCName(ByVal strDomainName As String) As String

   Dim strDCName As String
   Dim lngDCNPtr As Long
   Dim bytDNArray() As Byte
   Dim bytDCNArray(100) As Byte
   Dim lngResult As Long

   bytDNArray = strDomainName & vbNullChar
   
   'Lookup the Primary Domain Controller
   lngResult = NetGetDCName(0&, bytDNArray(0), lngDCNPtr)
   If lngResult <> 0 Then
      MsgBox "Error: " & lngResult
      Exit Function
   End If
   lstrcpyW bytDCNArray(0), lngDCNPtr
   lngResult = NetApiBufferFree(lngDCNPtr)
   strDCName = bytDCNArray()
   GetPDCName = Left(strDCName, InStr(strDCName, Chr(0)) - 1)

End Function

Public Function DisplaySystemError(ByVal lngErrCode As Long) As String
    
    Dim strMsg As String
    Dim strRtrnCode As String
    Dim lngFlags As Long
    Dim lngModule As Long
    Dim lngRet As Long

    lngModule = 0
    strRtrnCode = Space$(256)
    lngFlags = FORMAT_MESSAGE_FROM_SYSTEM
    
    'If lngRet is in the network range, load the message source
    If (lngErrCode >= NERR_BASE And lngErrCode <= MAX_NERR) Then
        lngModule = LoadLibraryEx("netmsg.dll", 0&, _
                  LOAD_LIBRARY_AS_DATAFILE)
    
        If (lngModule <> 0) Then
            lngFlags = lngFlags Or FORMAT_MESSAGE_FROM_HMODULE
        End If
    End If
    
    ' Call FormatMessage() to allow for message text to be acquired
    ' from the system or the supplied module handle.
    lngRet = FormatMessage(lngFlags, lngModule, lngErrCode, 0&, strRtrnCode, 256&, 0&)
    If lngRet = 0 Then
       strMsg = Err.LastDllError
    End If
    
    'If message source was loaded, unload it.
    If (lngModule <> 0) Then
        FreeLibrary (lngModule)
    End If

    If strMsg = "" Then
        strMsg = "ERROR: " & lngErrCode & " - " & strRtrnCode
    End If
    
    DisplaySystemError = strMsg

End Function











