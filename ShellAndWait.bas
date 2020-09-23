Attribute VB_Name = "ShellAndWait"
' Shell and Wait module from PSCODE.COM

Private Const INFINITE = &HFFFFFFFF
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const WAIT_TIMEOUT = &H102&



Private Type STARTUPINFO
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


Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessByNum Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes _
    As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags _
    As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As _
    STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

    


Public Function LaunchAppSynchronous(strExecutablePathAndName As String, Optional CommandLine As String, Optional Hidden As Boolean) As Boolean
    
    'Launches an executable by starting it's process
    'then waits for the execution to complete.
    '
    'INPUT: The executables full path and name.
    'RETURN: True upon termination if successful, false if not.
    
    Dim lngResponse As Long
    Dim typStartUpInfo As STARTUPINFO
    Dim typProcessInfo As PROCESS_INFORMATION
    LaunchAppSynchronous = False


    With typStartUpInfo
        .cb = Len(typStartUpInfo)
        .lpReserved = vbNullString
        .lpDesktop = vbNullString
        .lpTitle = vbNullString
        If Not Hidden Then
            .dwFlags = 0
        Else
            .dwFlags = &H1
            .wShowWindow = 0
        End If
    End With
    'Launch the application by creating a ne
    '     w process
    If CommandLine = "" Then
        lngResponse = CreateProcessByNum(strExecutablePathAndName, vbNullString, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, typStartUpInfo, typProcessInfo)
    Else
        lngResponse = CreateProcessByNum(vbNullString, strExecutablePathAndName & " " & CommandLine, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, typStartUpInfo, typProcessInfo)
    End If

    If lngResponse Then
        'Wait for the application to terminate b
        '     efore moving on
        Call WaitForTermination(typProcessInfo)
        LaunchAppSynchronous = True
    Else
        LaunchAppSynchronous = False
    End If
End Function


Private Sub WaitForTermination(typProcessInfo As PROCESS_INFORMATION)
    'This wait routine allows other applicat
    '     ion events
    'to be processed while waiting for the p
    '     rocess to
    'complete.
    Dim lngResponse As Long
    'Let the process initialize
    Call WaitForInputIdle(typProcessInfo.hProcess, INFINITE)
    'We don't need the thread handle so get
    '     rid of it
    Call CloseHandle(typProcessInfo.hThread)
    'Wait for the application to end


    Do
        lngResponse = WaitForSingleObject(typProcessInfo.hProcess, 0)


        If lngResponse <> WAIT_TIMEOUT Then
            'No timeout, app is terminated
            Exit Do
        End If


        DoEvents
        Loop While True
        'Kill the last handle of the process
        Call CloseHandle(typProcessInfo.hProcess)
    End Sub

