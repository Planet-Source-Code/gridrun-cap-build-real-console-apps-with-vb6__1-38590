VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select Target Subsystem"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CheckBox chkLinklog 
         Caption         =   "Create linklog.txt"
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   810
         Width           =   1500
      End
      Begin VB.CheckBox chkNoClean 
         Alignment       =   1  'Right Justify
         Caption         =   "Preserve .OBJ files after compiling"
         Height          =   375
         Left            =   1935
         TabIndex        =   5
         Top             =   810
         Width           =   2730
      End
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "Abort"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Console"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Windows"
         Default         =   -1  'True
         Height          =   375
         Left            =   3330
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Console Application Linker Proxy   (c) 2002 by gridrun [TNC]"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1350
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LinkerName As String


Private Sub Command1_Click()
    ' build for windows subsystem
    LockAllControls
    LogWrite "build for windows subsystem @ " & Now & "  -  command line: " & Command$ & "  -  OutPath: " & GetOutPath
    LaunchLinker Command$
    LogWrite "build done @ " & Now
    Unload Me
End Sub

Private Sub Command2_Click()
    ' build for console subsystem
    LockAllControls
    Dim CmdLine As String
    CmdLine = Replace(Command$, "/SUBSYSTEM:WINDOWS,4.0", "/SUBSYSTEM:CONSOLE")
    CmdLine = CmdLine & " /FORCE:UNRESOLVED"
    LogWrite "build for console subsystem @ " & Now & "  -  command line: " & CmdLine & "  -  OutPath: " & GetOutPath
    LaunchLinker CmdLine
    LogWrite "build done @ " & Now
    Unload Me
End Sub

Private Sub Command3_Click()
    ' abort build
    LockAllControls
    LogWrite "build aborted @ " & Now
    If chkNoClean.Value = 0 Then WipeObjFiles
    Unload Me
End Sub

Private Sub LaunchLinker(CommandLine As String)
    ' Shell actual LINK.EXE (LINK.ORIG.EXE by default)
    Dim LinkPath As String: LinkPath = AddSlash(App.Path) & LinkerName
    If Dir(LinkPath) = "" Then
        MsgBox "VB Linker (" & LinkerName & ") not found! Aborting.", vbCritical Or vbOKOnly, "CRITICAL ERROR"
        Unload Me
        Exit Sub
    End If
    If chkNoClean.Value = 1 Then
        LaunchAppSynchronous "command.com", "/c attrib +r " & AddSlash(GetOutPath) & "*.obj", True
        LaunchAppSynchronous LinkPath, CommandLine, True
        Shell "command.com /c attrib -r " & AddSlash(GetOutPath) & "*.obj", vbHide
    Else
        LaunchAppSynchronous "command.com", "/c attrib -r " & AddSlash(GetOutPath) & "*.obj", True
        LaunchAppSynchronous LinkPath, CommandLine, True
        WipeObjFiles
    End If
End Sub

Private Function GetOutPath() As String
    ' Extract Output File Path from Command Line Args
    Dim OutStart As Long: Dim OutEnd As Long
    OutStart = InStr(1, Command$, "/OUT:" & Chr(34)) + 1
    If OutStart = 0 Then Exit Function
    OutStart = OutStart + 5
    OutEnd = InStr(OutStart, Command$, Chr(34))
    If OutEnd = 0 Or OutEnd < OutStart Then Exit Function
    Dim OutPath As String
    OutPath = Mid(Command$, OutStart, OutEnd - OutStart)
    Dim FileNameStart As Long
    FileNameStart = InStrRev(OutPath, "\", -1)
    GetOutPath = Left$(OutPath, FileNameStart - 1)
End Function

Private Sub WipeObjFiles()
    ' clean up .OBJ files
    If chkLinklog.Value = 0 Then Exit Sub
    On Error Resume Next
    Kill AddSlash(GetOutPath) & "*.obj"
End Sub

Private Function AddSlash(Path As String) As String
    ' Helper to make sure a path ends with a backslash (shuld better be called AddBackSlash() or something, but heck, who cares? :P)
    If Right(Path, 1) = "\" Then
        AddSlash = Path
    Else
        AddSlash = Path & "\"
    End If
End Function

Private Function LogWrite(Line As String)
    ' write to linklog.txt
    On Error Resume Next
    If chkLinklog.Value = 0 Then Exit Function
    Dim fH As Integer: fH = FreeFile
    Open "linklog.txt" For Append As #fH
    Print #fH, Line
    Close #fH
End Function

Private Sub LockAllControls()
    ' Disable all input controls on the form
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    chkNoClean.Enabled = False
    chkLinklog.Enabled = False
End Sub

Private Sub Form_Load()
    ' init
    If Command$ = "" Then
        ' error msg
        Dim Prompt As String
        Prompt = "This application cannot be used directly, it must be launched by the VB IDE!" & vbCrLf & "(c) 2002 by gridrun [TNC]"
        MsgBox Prompt, vbExclamation Or vbOKOnly, "Console Application Linker Proxy"
        Unload Me
    Else
        ' load settings
        If GetSetting(App.EXEName, "Config", "LINKLOG", "1") = "1" Then
            chkLinklog.Value = 1
        Else
            chkLinklog.Value = 0
        End If
        If GetSetting(App.EXEName, "Config", "PRESERVEOBJ", "0") = "1" Then
            chkNoClean.Value = 1
        Else
            chkNoClean.Value = 0
        End If
        LinkerName = GetSetting(App.EXEName, "Config", "LINKER", "LINK.ORIG.EXE")
        LogWrite "-----------------------"
        LogWrite "invoked @ " & Now & "  -  Command Line: " & Command$
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' save settings
    If chkLinklog.Value = 1 Then
        SaveSetting App.EXEName, "Config", "LINKLOG", "1"
    Else
        SaveSetting App.EXEName, "Config", "LINKLOG", "0"
    End If
    If chkNoClean.Value = 1 Then
        SaveSetting App.EXEName, "Config", "PRESERVEOBJ", "1"
    Else
        SaveSetting App.EXEName, "Config", "PRESERVEOBJ", "0"
    End If
    Unload frmMain
    Set frmMain = Nothing
End Sub
