Attribute VB_Name = "modWaveIn"
Option Explicit

' DEBUGING
Public Const Debuging As Boolean = False


Enum OutputMode
    WAVE = 0
    MP3 = 1
End Enum

Enum eMP3_TYPE
    CBR = 0
    ABR = 1
    vbr = 2
End Enum

Enum eVBR_Routine
    New_Routine = 0
    Old_Routine = 1
End Enum

Public Type tMP3
    MP3_Type As eMP3_TYPE
    
    VBR_MinBitrate As Integer
    VBR_MaxBitrate As Integer
    VBR_Quality As Integer
    VBR_Routine As eVBR_Routine
    
    ABR_AvgBitrate As Integer
    
    CBR_Bitrate As Integer
    
    LAME As String
End Type


Public INI_FILE As String

Private Const RECORD As Long = 100
Private Const MONITOR As Long = 10

Private Const GMEM_FIXED As Long = &H0

Private Const CALLBACK_WINDOW As Long = &H10000

Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING

Private Const WAVE_FORMAT_PCM As Long = 1

Private Const MM_WIM_CLOSE As Long = &H3BF
Private Const MM_WIM_DATA As Long = &H3C0
Private Const MM_WIM_OPEN As Long = &H3BE
Private Const WIM_CLOSE As Long = MM_WIM_CLOSE
Private Const WIM_DATA As Long = MM_WIM_DATA
Private Const WIM_OPEN As Long = MM_WIM_OPEN
Private Const WHDR_DONE As Long = &H1
Private Const PROCESS_TERMINATE As Long = (&H1)


Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type


Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Declare Function TerminateProcess Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByVal uExitCode As Long) As Long

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" ( _
     ByRef Destination As Any, _
     ByVal Length As Long)

Private Declare Function CloseHandle Lib "kernel32.dll" ( _
     ByVal hObject As Long) As Long

Private Const MB_ICONHAND As Long = &H10&
Private Const MB_ICONERROR As Long = MB_ICONHAND

Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" ( _
     ByVal hwnd As Long, _
     ByVal lpText As String, _
     ByVal lpCaption As String, _
     ByVal wType As Long) As Long


Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5


Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const STD_INPUT_HANDLE As Long = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&

Private Declare Function AllocConsole Lib "kernel32.dll" () As Long
Private Declare Function FreeConsole Lib "kernel32.dll" () As Long
Private Declare Function GetStdHandle Lib "kernel32.dll" ( _
     ByVal nStdHandle As Long) As Long


Private Declare Function GetExitCodeProcess Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByRef lpExitCode As Long) As Long

Private Declare Function CreatePipe Lib "kernel32" ( _
    ByRef phReadPipe As Long, _
    ByRef phWritePipe As Long, _
    ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Private Declare Sub GetStartupInfoA Lib "kernel32" ( _
    ByRef lpInfo As STARTUPINFO)

Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByRef lpProcessAttributes As Any, _
    ByRef lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByRef lpEnvironment As Any, _
    ByVal lpCurrentDriectory As String, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
    

Private Declare Function WriteFile Lib "kernel32.dll" ( _
     ByVal hFile As Long, _
     ByVal lpBuffer As String, _
     ByVal nNumberOfBytesToWrite As Long, _
     ByRef lpNumberOfBytesWritten As Long, _
     ByRef lpOverlapped As Any) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByRef lpOverlapped As Any) As Long
  
Private Declare Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hReadPipe As Long, _
    ByRef lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, _
    ByRef lpTotalBytesAvail As Long, _
    ByRef lpBytesLeftThisMessage As Long) As Long


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMiliseconds As Long)

Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" ( _
     ByVal sz As String, _
     ByVal uFlags As Long) As Long

Private Type tHeader
    RIFF As Long            ' "RIFF"
    LenR As Long            ' size of following segment
    WAVE As Long            ' "WAVE"
    fmt As Long             ' "fmt
    FormatSize As Long      ' chunksize
    format As WAVEFORMATEX  ' audio format
    data As Long            ' "data"
    DataLength As Long      ' length of datastream
End Type



Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByRef Destination As Any, _
     ByRef Source As Any, _
     ByVal Length As Long)

Public Declare Sub CopyAudioMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByRef Destination As Any, _
     ByVal Source As Long, _
     ByVal Length As Long)


Private Declare Function GlobalAlloc Lib "kernel32.dll" ( _
     ByVal wFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" ( _
     ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" ( _
     ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
     ByVal hMem As Long) As Long


Private Const GWL_WNDPROC As Long = -4

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
     ByVal lpPrevWndFunc As Long, _
     ByVal hwnd As Long, _
     ByVal msg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
     ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long




Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" ( _
     ByVal err As Long, _
     ByVal lpText As String, _
     ByVal uSize As Long) As Long


Private Declare Function waveInReset Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInOpen Lib "winmm.dll" ( _
     ByRef lphWaveIn As Long, _
     ByVal uDeviceID As Long, _
     ByRef lpFormat As WAVEFORMATEX, _
     ByVal dwCallback As Long, _
     ByVal dwInstance As Long, _
     ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long

Private hStdinWrite As Long
Private hStdoutWrite As Long
Private hStdinRead As Long
Private hStdoutRead As Long

Private pi As PROCESS_INFORMATION

Private Const BUFFERS As Integer = 4
Private Const BUFFERS_MONITOR As Integer = 2
Private Const BUFFERSIZE_MONITOR As Integer = 8192
Private BUFFERSIZE As Long
'Private Const BUFFERSIZE As Long = 8192

Dim hWaveIn As Long
Public hWaveIn_Monitor As Long
Dim ret As Long
Dim format As WAVEFORMATEX
Dim hMem(BUFFERS) As Long
Dim hmem_monitor(BUFFERS_MONITOR) As Long
Dim hdr(BUFFERS) As WAVEHDR
Dim hdr_monitor(BUFFERS_MONITOR) As WAVEHDR
Dim lpPrevWndFunc As Long
Dim hwnd As Long
Dim num As Integer
Public msg As String * 255
Dim pos As Long
Dim pos_mp3 As Long
Dim bHeaderWritten As Boolean
Public bRecording As Boolean
Public bMonitoring As Boolean
Dim bPaused As Boolean
Dim bPaused_Monitoring As Boolean
Dim OutputFile As String
Public OutMode As OutputMode
Public MP3_Settings As tMP3

Dim curBuffer As Long

' =============
' Subclassing function
' =============
Sub Hook(bHook As Boolean)
    ' Save hWnd of main form for waveOutOpen purposes => see PrepareRecording sub
    hwnd = frmMain.hwnd
    'Exit Sub
    
    ' Prevent double hooking
    If lpPrevWndFunc <> 0 And bHook Then
        MessageBox 0&, "Double-hooking not permited!", "ERROR!", MB_ICONERROR
        Exit Sub
    End If
    
    Logging "[Hook] About to hook/unhook frmMain"
    
    ' bHook chooses whether we should subclass the form or restore previous Callback
    ' function
    If bHook Then
        lpPrevWndFunc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf CallbackProc)
    Else
        lpPrevWndFunc = 0
        SetWindowLong hwnd, GWL_WNDPROC, lpPrevWndFunc
    End If
End Sub

' ================
' Pausing function
' ----------------
' It's a toggling function
' ================
Sub Pause()
    Dim i As Integer
    
    bPaused = Not bPaused
    
    If bPaused Then
        ' Order MM subsystem to stop sending recorded data
        waveInStop hWaveIn
    Else
        ' We have to add buffers to MM subsystem, otherwise it won't send any new data
        For i = 0 To BUFFERS
            waveInAddBuffer hWaveIn, hdr(i), Len(hdr(i))
        Next
        
        ' Finally start again recording
        waveInStart hWaveIn
    End If
End Sub

' internal function: delete recording file (obviously, used after recording has stopped)
Private Sub DeleteFile()
    Close #num
    Kill OutputFile
End Sub

' Resets monitoring
' =================
' Sometimes does monitoring to 'stuck'
' In most cases this solved it
'
Public Sub ResetMonitoring()
    waveInReset hWaveIn_Monitor
End Sub

' ==================
' Stop monitoring
' ------------------
' this stops secondary recording for monitoring
' It should be called at program quit. To pause use PauseMonitoring
' ==================
Public Sub StopMonitoring()
    Dim i As Integer
    
    ' Signal that recording for monitoring has stopped (no longer available)
    bMonitoring = False
    
    ' Stops monitoring
    ret = waveInStop(hWaveIn_Monitor)
    
    ' Reset it
    ret = waveInReset(hWaveIn_Monitor)
    
    ' Unprepare buffers
    For i = 0 To BUFFERS_MONITOR
        waveInUnprepareHeader hWaveIn_Monitor, hdr_monitor(i), Len(hdr_monitor(i))
        GlobalUnlock hdr_monitor(i).lpData
        GlobalFree hdr_monitor(i).lpData
    Next i
    
    ' Close hWaveIn_Monitor
    ret = waveInClose(hWaveIn_Monitor)
    
    ' Clear handle
    hWaveIn_Monitor = 0
End Sub

' ===============
' Stops recording
' ---------------
' this stops main recording
'================
Sub StopRec()
    Dim i As Integer
    
    ' if not recording, exit sub
    If Not bRecording Then Exit Sub
    
    ' stop sending data
    ret = waveInStop(hWaveIn)
    
    ' reset hWaveIn
    ret = waveInReset(hWaveIn)
    
    ' Unprepare buffers
    For i = 0 To BUFFERS
        waveInUnprepareHeader hWaveIn, hdr(i), Len(hdr(i))
        GlobalUnlock hdr(i).lpData
        GlobalFree hdr(i).lpData
    Next i
    
    ' Close device
    ret = waveInClose(hWaveIn)
    
    ' Write header only if its not already written AND output _MUST_ be WAVE (obviously)
    If (bHeaderWritten = False) And (OutMode = WAVE) Then WriteWAVHeader
    
    ' closes recorded file
    Close #num
    
    ' Signal that recording has stopped
    bRecording = False
    
    ' If recording to MP3 stop (kill) encoder
    ' killing lame.exe is because Lame just can't know _when_ we want to stop it
    If OutMode = MP3 Then StopEncoder
    
    ' Clear handle
    hWaveIn = 0
End Sub

' =====================
' Pause monitoring : toggling function
' =====================
Public Sub PauseMonitoring()
    Dim i As Integer
    
    ' Toggle monitoring signal "Paused"
    bPaused_Monitoring = Not bPaused_Monitoring
    
    If bPaused_Monitoring Then
        ' stop sending data
        waveInStop hWaveIn_Monitor
    Else
        ' We have to add buffers or we won't get any data
        For i = 0 To BUFFERS_MONITOR
            waveInAddBuffer hWaveIn_Monitor, hdr_monitor(i), Len(hdr_monitor(i))
        Next
        
        ' Reset it
        waveInReset hWaveIn_Monitor
        
        ' Start monitoring (hWaveIn_Monitor)
        waveInStart hWaveIn_Monitor
    End If
End Sub

' =======
' Prepares monitoring
' =======
Public Sub PrepareMonitoring()
    Dim formatMonitor As WAVEFORMATEX
    Dim i As Integer
    
    ' Set constant parameters for monitoring
    With formatMonitor
        .wFormatTag = 1
        .nChannels = 2
        .wBitsPerSample = 16
        .nSamplesPerSec = 44100
        .nBlockAlign = .nChannels * .wBitsPerSample / 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With
    
    ' open recording (hWaveIn_Monitor) with format (formatMonitor), send callbacks
    ' (CALLBACK_WINDOW) to main form (frmMain.hwnd)
    ret = waveInOpen(hWaveIn_Monitor, 0, formatMonitor, hwnd, 0&, CALLBACK_WINDOW)
    If ret <> 0 Then
        ' Was there an error? Catch it and display
        waveInGetErrorText ret, msg, Len(msg)
        MessageBox 0&, Trim(msg), App.Title, MB_ICONERROR
        ' quit
        Exit Sub
    End If
    
    ' Preparing buffers for monitoring
    For i = 0 To BUFFERS_MONITOR
        ' Allocate space buffer (size BUFFERSIZE_MONITOR)
        ' Notice GMEM_FIXED, it's essential, because otherwise MM system won't be able
        ' to write into it.
        hmem_monitor(i) = GlobalAlloc(GMEM_FIXED, BUFFERSIZE_MONITOR)
        
        With hdr_monitor(i)
            .lpData = GlobalLock(hmem_monitor(i)) ' Lock buffer (I don't know exactly why...)
            .dwBufferLength = BUFFERSIZE_MONITOR  ' Set buffer length
            .dwFlags = 0                          ' no flags
            .dwLoops = 0                          ' no loops
            .dwUser = CLng(i) + MONITOR           ' Here we put _ID_ of buffer
            
            ' Notice MONITOR constant (up). It's for callback function to recognize
            ' whether it's montoring or recording data
        End With
    Next
    
    ' Prepare buffers + Add them to MM system
    For i = 0 To BUFFERS_MONITOR
        ret = waveInPrepareHeader(hWaveIn_Monitor, hdr_monitor(i), Len(hdr_monitor(i)))
        
        ret = waveInAddBuffer(hWaveIn_Monitor, hdr_monitor(i), Len(hdr_monitor(i)))
    Next
    
    ' Signal that recording for monitoring has started
    bMonitoring = True
    
    ' Start recording (into SECONDARY buffer << for monitoring)
    ret = waveInStart(hWaveIn_Monitor)
End Sub

' =============
' Prepares recording: MAIN RECORDING
' =============
Sub PrepareRecording( _
    ByVal iChannels As Integer, _
    ByVal iBits As Integer, _
    ByVal lFrequency As Long, _
    ByVal OutputMode1 As OutputMode, _
    Optional ByVal sFile As String)
    
    Dim i As Integer
    
    ' Signal that we're recording now
    bRecording = True
    
    ' Zero current positions
    pos = 0
    pos_mp3 = 0
    
    ' Enable access other function in module to output file
    OutputFile = sFile
    
    ' Creates unique number for file.
    ' Using CreateFile + WriteFile would be confusing and could create more bugs, which
    ' could lead to another bugs
    num = FreeFile
    
    ' Standard prepares for file writing
    If Dir(sFile) <> "" Then Kill sFile
    Open sFile For Binary As #num
    
    ' move after header (= first 44 bytes of WAVE file)
    pos = 45
    bHeaderWritten = False          ' reset indicator whether the header is written
    
    ' Set format to parameters, which has user chosen
    With format
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = iChannels
        .wBitsPerSample = iBits
        .nSamplesPerSec = lFrequency
        .nBlockAlign = .nChannels * .wBitsPerSample / 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
        
        ' This buffer calculation was fine for WAVE recording. But problems have risen
        ' up with MP3 direct recording (like lockups after few recorded frames).
        ' Now the better calculation is little below.
        'BUFFERSIZE = .nSamplesPerSec * .nBlockAlign * .nChannels * 0.1
        
        ' Calculation proper for MP3 recording
        BUFFERSIZE = (.nSamplesPerSec * .wBitsPerSample / 8) * 0.1
        BUFFERSIZE = BUFFERSIZE - (BUFFERSIZE Mod .nBlockAlign)
    End With
    
    ' Open recording handle
    ret = waveInOpen(hWaveIn, 0, format, hwnd, 0&, CALLBACK_WINDOW)
    If ret <> 0 Then
        ' If there's an error, display it
        waveInGetErrorText ret, msg, Len(msg)
        MessageBox 0&, Trim(msg), App.Title, MB_ICONERROR
        Exit Sub
    End If
    
    ' Prepare buffers
    For i = 0 To BUFFERS
        ' Allocate space buffer (size BUFFERSIZE_MONITOR)
        ' Notice GMEM_FIXED, it's essential, because otherwise MM system won't be able
        ' to write into it.
        hMem(i) = GlobalAlloc(GMEM_FIXED, BUFFERSIZE)
        
        With hdr(i)
            .lpData = GlobalLock(hMem(i))
            .dwBufferLength = BUFFERSIZE
            .dwFlags = 0
            .dwLoops = 0
            .dwUser = CLng(i) + RECORD
            ' Notice RECORD constant (up). It's for callback function to recognize
            ' whether it's montoring or recording data
        End With
    Next
    
    ' Prepare buffers + Add them to MM system
    For i = 0 To BUFFERS
        ret = waveInPrepareHeader(hWaveIn, hdr(i), Len(hdr(i)))
        
        ret = waveInAddBuffer(hWaveIn, hdr(i), Len(hdr(i)))
    Next
    
    ' Start recording
    ret = waveInStart(hWaveIn)
    OutMode = OutputMode1     ' Save output mode for other module functions
    
    If OutMode = MP3 Then
        PrepareEncoder      ' We need to start LAME, as output is now MP3
    End If
End Sub

' ===========
' Complicated preparations for LAME
' ===========

Private Sub PrepareEncoder()
    Dim PID As Long
    Dim pa As SECURITY_ATTRIBUTES
    Dim pra As SECURITY_ATTRIBUTES
    Dim tra As SECURITY_ATTRIBUTES
    Dim sui As STARTUPINFO
    
    Dim cmdLine As String
    
    Dim hStderr As Long
    
    Dim ret As Long
    Dim LameExe As String
    Dim temp As String
    Dim bAvail As Long
    Dim bRead As Long
    Dim lExitCode As Long
    
    pa.nLength = Len(pa)
    pa.bInheritHandle = 1
    
    pra.nLength = Len(pra)
    tra.nLength = Len(tra)
    
    ' Create pipe for LAME output (one write-only for LAME, and one read-only for us)
    CreatePipe hStdoutRead, hStdoutWrite, pa, 0
    If hStdoutRead = 0 Then Exit Sub
    
    ' Create pipe for LAME input (one read-only for LAME, and one write-only for us)
    CreatePipe hStdinRead, hStdinWrite, pa, 0
    If hStdinRead = 0 Then Exit Sub
    
    ' Get startup info for audio_recorder.exe (VB6.exe)
    sui.cb = Len(sui)
    GetStartupInfoA sui
    sui.hStdOutput = hStdoutWrite   ' this is essential, to set those handles
    sui.hStdInput = hStdinRead      ' so we would be able to send RAW audio to LAME
                                    ' and LAME would send us MP3 data
    sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES ' order to use those
                                                            ' handles and to hide window
    sui.wShowWindow = SW_HIDE       ' this is essential to hide ugly console window
    
    ' complicated command line building (for LAME)
    With MP3_Settings
        ' Get Lame path
        If .LAME = "" Then
            LameExe = AppPath & "lame.exe"
        Else
            If Right(.LAME, 1) = "\" Then
                LameExe = .LAME & "lame.exe"
            Else
                LameExe = .LAME & "lame.exe"
            End If
        End If
        
        ' check whether lame.exe exists there
        If Dir(LameExe) = "" Then
            StopRec
            DeleteFile
            MessageBox 0&, "ERROR!!!!!!!!!!!!" & vbCrLf & vbCrLf & LameExe & _
                vbCrLf & vbCrLf & "Doesn't exist!!!!!", App.Title, MB_ICONERROR
            Exit Sub
        End If
        
        ' Build cmd. line according to user selection of encoding mode
        If .MP3_Type = CBR Then
            cmdLine = " -b " & CStr(.CBR_Bitrate)   ' only bitrate is needed here
        ElseIf .MP3_Type = ABR Then
            cmdLine = " --abr " & CStr(.ABR_AvgBitrate)  ' again only bitrate
        ElseIf .MP3_Type = vbr Then
            If .VBR_Routine = New_Routine Then          ' User wants to use old routine
                cmdLine = " --vbr-new"
            ElseIf .VBR_Routine = Old_Routine Then
                cmdLine = " --vbr-old"                  ' for VBR or new one?
            End If
            
            ' ensure that VBR quality is in boundaries 0..9
            If .VBR_Quality > 9 Then .VBR_Quality = 9
            If .VBR_Quality < 0 Then .VBR_Quality = 0
            
            cmdLine = cmdLine & " -V " & CStr(.VBR_Quality)     ' add VBR quality to
                                                                ' command line
            cmdLine = cmdLine & " -b " & CStr(.VBR_MinBitrate)  ' min. bitrate for VBR
            cmdLine = cmdLine & " -B " & CStr(.VBR_MaxBitrate)  ' max. bitrate for VBR
            
            cmdLine = cmdLine & " -m "                ' choose between mono and stereo
            If format.nChannels = 1 Then              ' according to recording settings
                cmdLine = cmdLine & "m"     ' mono
            Else
                cmdLine = cmdLine & "j"     ' joint-stereo
            End If
        End If
        
        ' other essentials: indicate LAME sampling frequency, that we'll send raw data,
        ' and to swap bytes (I hate Big/Little endian mess :D )
        cmdLine = " -s " & Replace(CStr(format.nSamplesPerSec / 1000), ",", ".") & _
            " -r -x --bit-width" & format.wBitsPerSample & cmdLine & " - -"
        
        cmdLine = LameExe & cmdLine     ' And finally add Lame.exe to the command line
        frmMain.DebugIt cmdLine         ' show to user command line
    End With
    
    ' Create process
    PID = CreateProcessA(vbNullString, cmdLine, pra, tra, 1, 0, Null, vbNullString, sui, pi)
    Sleep 20                                    ' we have to wait otherwise
                                                ' GetExitCodeProcess will hang app.
    GetExitCodeProcess pi.hProcess, lExitCode
    
    If PID = 0 Or lExitCode <> STILL_ACTIVE Then ' Is lame still running
        
        ' Has Lame sent us something, before it died? ;D
        If PeekNamedPipe(hStdoutRead, ByVal 0, 0, ByVal 0, bAvail, ByVal 0) Then
            DoEvents
            
            ' Read the message...
            If bAvail Then
                temp = String(bAvail, 0)
                ReadFile hStdoutRead, temp, bAvail, bRead, ByVal 0
                CloseHandle hStdoutWrite   ' this turns the pipe around
            End If
            
            ' ... and show to user
            frmMain.DebugIt temp
        Else
            ' Display error message
            MessageBox 0&, "Error running encoder!", frmMain.Caption, MB_ICONERROR
        End If
        
        ' Interrupt recording
        DeleteFile
        StopRec
        Exit Sub
    End If
End Sub

Private Sub StopEncoder()
    Dim lExitCode As Long
    
    CloseHandle hStdoutWrite
    
    CloseHandle hStdoutRead
    CloseHandle hStdinRead
    CloseHandle hStdinWrite
    
    GetExitCodeProcess pi.hProcess, lExitCode
    
    TerminateProcess pi.hProcess, lExitCode
    
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
End Sub

' ======================================================================================
' MAIN FUNCTION where the Big Show goes on (e.g. detecting whether we've got some data)
' It's a risky callback function of subclassing for frmMain.  If an error is created
' here, then "we're finished". Boom. Crash.
' ======================================================================================

Public Function CallbackProc(ByVal hw As Long, ByVal uMsg As Integer, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long
    Dim temp() As Byte
    Dim i As Integer
    
    On Error Resume Next
    
    ' Detect according uMsg whether it's something interesting data for us
    If (uMsg = WIM_DATA) Then
        
        ' This is primary recording buffer
        If wavhdr.dwUser - RECORD >= 0 And bRecording = True And bPaused = False Then
            
            ' Check whether the current buffer is filled with new audio
            If (wavhdr.dwFlags And WHDR_DONE) Then
                
                Logging "[CallbackProc] Retrieving recorded data from wavhdr.lpData"
                
                ' Allocate data for buffer, which is used to Write Data to disk
                ' (either as raw WAVE or encode MP3)
                ReDim temp(wavhdr.dwBytesRecorded - 1)
                
                ' Copy buffer from wavhdr (WAVEHDR struct) to our (temp)
                CopyMemory temp(0), ByVal wavhdr.lpData, wavhdr.dwBytesRecorded
                
                Logging "[CallbackProc] About to write data to Output file (" & CStr(OutMode) & ")"
                
                ' Convert data from bytes to string
                WriteData StrConv(temp, vbUnicode)
                
            End If
            
            Logging "[CallbackProc] About to Add buffers to mmsystem"
            
            ' Go through buffers, and those, which aren't filled with new audio,
            ' add them to MM system (waveInAddBuffer)
            For i = 0 To (BUFFERS)
                If Not (hdr(i).dwFlags And WHDR_DONE) Then
                    Logging "[CallbackProc] Adding buffer with dwUser = " & hdr(i).dwUser
                    ret = waveInAddBuffer(hWaveIn, hdr(i), Len(hdr(i)))
                End If
            Next
            
            ' If there's an error, this is the way, which won't hurt our application
            ' because of subclassing (put it in public variable and Timer in frmMain
            ' will add it as Debug information (DebugIt function)
            If err Then
                msg = err.Description
                err.Clear
            End If
            
            
        ' Just monitoring...… (e.g. secondary recording buffer)
        ElseIf bMonitoring = True And bPaused_Monitoring = False Then
            
            ' Put in curBuffer ID of current buffer. Actual peak calculation is
            ' called from Timer in frmMain. From here it might hurt, because of
            ' subclassing, although now it shouldn't either.
            ' =================================================================
            If (wavhdr.dwFlags And WHDR_DONE) Then
                curBuffer = wavhdr.dwUser - MONITOR
            End If
            
            Logging "[CallbackProc] About to Add buffers for MONITORING"
            
            ' Go through buffers, and those, which aren't filled with new audio,
            ' add them to MM system (waveInAddBuffer)
            For i = 0 To (BUFFERS_MONITOR)
                If Not (hdr_monitor(i).dwFlags And WHDR_DONE) Then
                    Logging "[CallbackProc] Adding buffer (FOR MONITORING) with dwUser" & hdr(i).dwUser
                    
                    ret = waveInAddBuffer(hWaveIn_Monitor, hdr_monitor(i), _
                        Len(hdr_monitor(i)))
                End If
            Next i
            
            ' If there's an error, this is the way, which won't hurt our application
            ' because of subclassing (put it in public variable and Timer in frmMain
            ' will add it as Debug information (DebugIt function)
            If err Then
                msg = err.Description
                err.Clear
            End If
        End If
    End If
    
    On Error GoTo 0
    
    ' Previous callback function will process it. Except of WOM_CLOSE, WOM_DONE
    ' messages, because the default Window Handler doesn't handle it.
    CallbackProc = CallWindowProc(lpPrevWndFunc, hw, uMsg, wParam, wavhdr)
End Function

' =====================================
' Encoding/Raw WAVE writing procedure
' Another big show goes on here...
' =====================================
Sub WriteData(ByRef data As String)
    ' WAVE
    Dim temp() As Byte
    
    ' MP3 (lame)
    Dim bAvail As Long    ' pipe bytes available (PeekNamedPipe)
    Dim bRead As Long     ' pipe bytes fetched   (ReadFile)
    Dim bWrite As Long
    Dim encodedData As String
    Dim BufLen As Long
    Dim start As Long
    
    Logging "[WriteData] About to write output data"
    
    ' WAVE => write, MP3 => encode WAVE => write
    If OutMode = MP3 Then
        Logging "[WriteData] OutMode = MP3, about to write data to LAME"
        
        DoEvents
        
        ' Write into LAME, it's gonna encode it for us
        BufLen = Len(data)
        ret = WriteFile(hStdinWrite, data, BufLen, bWrite, ByVal 0)
        
        Logging "[WriteData] about to Peek pipe for incoming data"
        
        ' Check whether the data is ready
        If PeekNamedPipe(hStdoutRead, ByVal 0, 0, ByVal 0, bAvail, ByVal 0) Then
            DoEvents
            
            ' Is really something there?
            If bAvail Then
                ' Allocate buffer
                encodedData = String(bAvail, 0)
                
                Logging "[WriteData] About to _READ_ from pipe"
                
                ' Read from pipe
                ReadFile hStdoutRead, encodedData, bAvail, bRead, ByVal 0
                
                ' Write to file
                Put #1, , encodedData
                CloseHandle hStdoutWrite   ' This turns the pipe around
            End If
        End If
        
        ' "pos" is number of raw WAVE data !recorded! (this doesn't get written)
        pos = pos + BufLen
        
        ' pos_mp3 is number of MP3 data !written!
        pos_mp3 = pos_mp3 + Len(encodedData)
    ElseIf OutMode = WAVE Then
        Logging "[WriteData] Writing Wave data"
        
        If pos = 0 Then pos = 1
        
        temp = StrConv(data, vbFromUnicode)
        
        ' Write to file
        Put #num, pos, temp
        
        ' Increment current position
        pos = pos + BUFFERSIZE
    End If
End Sub

' ===========================
' Important function for WAVE files
' ===========================
Sub WriteWAVHeader()
    Dim Header As tHeader
    Dim File() As Byte
    
    ' This is simplified version of writing WAVE header. Modified for this small audio
    ' recorder purposes.
    
    With Header
        .RIFF = mmioStringToFOURCC("RIFF", 0&)
        .WAVE = mmioStringToFOURCC("WAVE", 0&)
        .fmt = mmioStringToFOURCC("fmt ", 0&)
        .data = mmioStringToFOURCC("data", 0&)
        
        .format = format
        
        .FormatSize = 16
        .DataLength = LOF(1) - 44
        .LenR = Len(Header) + .DataLength - 8
        ReDim File(Len(Header))
        
        CopyMemory File(0), Header, Len(Header)
    End With
    
    ' Write it into file
    Put #num, 1, File
    
    ' Signal that we've written the Header
    bHeaderWritten = True
End Sub

' ==================================
' Some properties
' (unusual for modules, yes, I know)
' ==================================


Public Property Get BytesWritten() As Long
    If OutMode = WAVE Then
        BytesWritten = pos
    Else
        BytesWritten = pos_mp3
    End If
End Property

Public Property Get GetTime() As Long
    GetTime = pos \ format.nAvgBytesPerSec
End Property

Public Property Get SampleFrequency() As Long
    SampleFrequency = format.nSamplesPerSec
End Property

Public Property Get BlockAlign() As Long
    BlockAlign = format.nBlockAlign
End Property

Public Property Get BufferLength() As Long
    BufferLength = BUFFERSIZE
End Property

Public Property Get BufferData(ByVal lBuffer As Long) As Long
    BufferData = hdr(lBuffer).lpData
End Property

Public Property Get PeakMax() As Double
    PeakMax = 32767
End Property

' =========================
' Peak calculation function
' =========================
Public Function GetCurPeak(ByRef lLeft As Double, ByRef lRight As Double, Optional ByVal dB As Boolean) As Boolean
    Static buffer As Integer
    Static bFirst As Boolean
    
    Dim maxLeft As Double
    Dim maxRight As Double
    
    Dim tmpBuffer(87) As Integer
    Dim curLeft As Double
    Dim curRight As Double
    Dim i As Integer
    Dim size As Long
    
    Dim temp As String
    
    
    Logging "[GetCurPeak] Calculating peak"
    
    
    On Error Resume Next
    
    ' Some check logic, that it won't start, when it isn't supposed to start
    If (bFirst = False) And (buffer <> curBuffer) Then
        bFirst = True
    ElseIf (bFirst = False) And (buffer = 0) Then
        Exit Function
    End If
    
    ' Buffer iteration
    If curBuffer = 0 Then
        buffer = BUFFERS_MONITOR
    Else
        buffer = curBuffer - 1
    End If
    
    ' If buffer does screw up, we should know... and be ready for it
    If buffer > BUFFERS_MONITOR Then
        temp = "GetCurPeak: buffer = " & CStr(buffer) & "  max = " & CStr(BUFFERS_MONITOR)
        
        If buffer < MONITOR + BUFFERS_MONITOR Then
            buffer = buffer - MONITOR
            temp = temp & "  FIXING to " & buffer
            frmMain.DebugIt temp
        Else
            Exit Function
        End If
    End If
    
    Logging "[GetCurPeak] _Copying_ audio data from lpData buffer"
    
    ' Copy current audio into temporary buffer
    CopyAudioMemory tmpBuffer(0), hdr_monitor(buffer).lpData, 88
    
    For i = 0 To UBound(tmpBuffer) Step 2
        ' Peak calculation, Left channel
        curLeft = Abs(tmpBuffer(i))
        If curLeft > maxLeft Then maxLeft = curLeft
        
        ' Peak calculation, Right channel
        curRight = Abs(tmpBuffer(i + 1))
        If curRight > maxRight Then maxRight = curRight
    Next i
    
    ' lLeft and lRight were passed ByRef, send the peak values back
    lLeft = maxLeft
    lRight = maxRight
    
    ' Check whether user wants dB and not linear Peak
    If dB Then
        Logging "About to calculate dB"
        
        lLeft = Calc_dB(maxLeft)
        lRight = Calc_dB(maxRight)
    End If
    
    GetCurPeak = True
    
    On Error GoTo 0
End Function

Public Function Calc_dB(ByVal lCurPeak As Double) As Double
    ' see wikipedia for more about decibels
    ' this took a lot to find out
    If lCurPeak = 0 Then
        Calc_dB = -120
        Exit Function
    End If
    
    Calc_dB = 20 * (Log(lCurPeak / PeakMax))
End Function

Function AppPath(Optional Path As String) As String
    If Path = "" Then Path = App.Path
    
    If Right(Path, 1) = "\" Then
        AppPath = Path
    Else
        AppPath = Path & "\"
    End If
End Function

Sub Logging(ByVal StrTest As String)
    Dim numLog
    
    If Debuging = False Then Exit Sub
    
    numLog = FreeFile
    Open "C:\test.log" For Append As numLog
    
    Print #numLog, StrTest
    
    Close numLog
End Sub
