Attribute VB_Name = "basDOSCon"
Option Explicit


' basDOSCon.bas, v 1.51 - PUBLIC VERSION
'
'
' Real Console Applications  -  It IS possible!!
'
' brought to you by gridrun [TNC] in 2002
' thanks to udi shitrit for his GUI linker and to vbworld.com for theyre console API tutorial
'
' This module will not work unless u compile ur project with gridrun's cap.exe or a similar
' tool, to console subsytem. Once this is done, this module will give you the ability to
' ConRead() from console and ConWrite() or ConPrint() to it, and more.
'
' In contrast to most other console solutions (most of them do not even give a real console,
' ie they just open up theyr own console), this module will allow you to debug your console
' app in your VBE IDE! Read gridrun's console app tutorial for more information.
'
' Additionally, this module tries to use STDIO whenever possible, meaning that your console
' applications will be able to work in a redirected environment. An example would be a
' DOS output redirection: c:\folder\app.exe > textfile.txt
'
' However, you can modify this behaviour and force usage of console API by setting IOMode = 1
'
' further dox bout the console api can be found in msdn. check out
'   http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dllproc/conchar_34dr.asp
' for more info and function declarations, heh.
'
' shouts to d1cer, ampoz, spin, smokeyser, vortek and stupidcap @#@#!!
' msg gridrun@undernet wif bug reports and updated versions :P
'
'



' Helper API functions
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long


' Console API
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long


' File IO functions
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long


' Helper function constants & types
Const VER_PLATFORM_WIN32_NT = 2 'win nt,2000,XP
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

' Handle constants
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

' Color constants
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80

' For SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8

' For SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

' Handles
Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long

' This will be set by our detection algorithm to work around a problem on win9x with ReadFile()
Private OverrideSTDIN As Boolean


' True if this runs inside a DOS window, false in IDE
Public ConAttached As Boolean

' Select IO Mode behaviour:
' 0 = use STDIO (redirection supported)
' 1 = use ConAPI (extended control support, no redirection)
Public IOMethod As Integer

Public Sub ConAcquire()
    ' default IO mode is STDIO
    IOMethod = 1
    ' check for wintendo9x & co OS
    If Not IsWinNT Then OverrideSTDIN = True ' these have a problem with ReadFile(), do not use it
    ' try to allocate handles for stderr/stdout/stdin from *current* console
    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
    ConAttached = True  ' this will tell us if we have to call FreeConsole() or not upon exit
    ' did we get our handles, above??
    ' According to MSDN, GetStdHandle() returns -1 (INVALID_HANDLE_VALUE) if a handle is not available
    ' but we have found that this is not always the case, hence the values 4, 8 and 12
    If hConsoleIn < 4 Or hConsoleOut < 8 Or hConsoleErr < 12 Then
        ' we dont have one. create a new console window
        ' (thanks to this clever trick you can now debug your con app from your VB IDE. gimme a w00t w00t :P)
        AllocConsole
        DoEvents
        ConAttached = False
        ' get handles
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
    End If
End Sub

Public Sub ConRelease()
    CloseHandle hConsoleIn: CloseHandle hConsoleOut: CloseHandle hConsoleErr ' close handles
    If Not ConAttached Then FreeConsole ' free up if the console was allocated!
End Sub



' stdErr - Write Line to Error Channel
Public Function ConErr(szError As String)
    If IOMethod = 1 Then
        WriteConsole hConsoleErr, szError & vbCrLf, Len(szError) + 2, vbNull, vbNull
    Else
        Dim Result As Long
        WriteFile hConsoleErr, szError & vbCrLf, Len(szError) + 2, Result, ByVal 0&
    End If
End Function

' stdOut - Write Chars to Output Channel
Public Sub ConWrite(szOut As String)
    If IOMethod = 1 Then
        WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull
    Else
        Dim Result As Long
        WriteFile hConsoleOut, szOut, Len(szOut), Result, ByVal 0&
    End If
End Sub

' stdOut - Write Line to Output Channel
Public Sub ConPrint(Optional ByVal szOut As String)
    ConWrite szOut & vbCrLf
End Sub

' stdIn - Read Line from Input Channel (synchronous)
Public Function ConRead() As String
    Call SetConsoleMode(hConsoleIn, ENABLE_PROCESSED_INPUT Or ENABLE_LINE_INPUT Or ENABLE_LINE_INPUT)
    Dim sUserInput As String * 256  ' max line lenght = 256chars
    If (IOMethod = 1) Or (OverrideSTDIN = True) Then
        Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    Else
        Dim BytesRead As Long
        Call ReadFile(hConsoleIn, ByVal sUserInput, Len(sUserInput), BytesRead, vbNull)
    End If
    'Trim off the NULL characters and the CRLF.
    On Error Resume Next
    ConRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

' Helper function to determine if a Windows NT based OS is running
Private Function IsWinNT() As Boolean
    Dim osv As OSVERSIONINFOEX
    osv.dwOSVersionInfoSize = Len(osv)
    If GetVersionEx(osv) = 1 Then
        If osv.dwPlatformId = VER_PLATFORM_WIN32_NT Then IsWinNT = True
    End If
End Function



