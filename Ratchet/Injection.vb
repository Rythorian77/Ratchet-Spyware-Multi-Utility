Module Injection

#Region "1. Global Declarations"

    Public ProsH As Long
    Public Verify As Integer

#End Region

#Region "2. API Functions"

    Private Declare Function GetProcAddress Lib "kernel32" (hModule As Long,
                                                                lpProcName As String) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (lpModuleName As String) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (lpLibFileName As String) As Long
    Private Declare Function VirtualAllocEx Lib "kernel32" (hProcess As IntPtr,
                                                            lpAddress As Long,
                                                            dwSize As Long,
                                                            fAllocType As Long,
                                                            flProtect As Long) As Long
    Public Declare Function WriteProcessMemory Lib "kernel32" (hProcess As IntPtr,
                                                               lpBaseAddress As Long,
                                                               lpBuffer As String,
                                                               nSize As Long,
                                                               lpNumberOfBytesWritten As Long) As Long
    Private Declare Function CreateRemoteThread Lib "kernel32" (ProcessHandle As IntPtr,
                                                                lpThreadAttributes As Long,
                                                                dwStackSize As Long,
                                                                lpStartAddress As Long,
                                                                lpParameter As Long,
                                                                dwCreationFlags As Long,
                                                                lpThreadID As Long) As Long
    Public Declare Function FreeLibrary Lib "kernel32" (hLibModule As Long) As Long

#End Region

#Region "3. Inject | Eject Functions"

    'The Injection Function
    Public Function InjectDll(DllPath As String,
                                  ProsH As IntPtr)
        Dim DLLVirtLoc As Long, Inject As Long, LibAddress As Long
        Dim CreateThread As Long, ThreadID As Long
        Dim DllLength As Long
        'STEP 1 -  The easy part...Putting the it in the process' memory
        'Find a nice spot for your DLL to chill using VirtualAllocEx
        DllLength = Len(DllPath)
        MsgBox(DllLength)
        DLLVirtLoc = VirtualAllocEx(ProsH,
                                        0,
                                        DllLength,
                                        &H1000,
                                        &H4)
        'Inject the Dll into that spot
        Inject = WriteProcessMemory(ProsH,
                                        DLLVirtLoc,
                                        DllPath,
                                        DllLength,
                                        vbNull)
        'STEP 2 - Loading it in the process
        'This is where it gets a little interesting....
        'Just throwing our Dll into the process isnt going to do nothing unless you
        'Load it into the precess address using LoadLibrary.  The LoadLibrary function
        'maps the specified executable module into the address space of the
        'calling process.  You call LoadLibrary by using CreateRemoteThread to
        'create a thread(no ____) that runs in the address space of another process.
        'First we find the LoadLibrary function in kernel32.dll
        LibAddress = GetProcAddress(GetModuleHandle("kernel32.dll"),
                                        "LoadLibraryA")
        'Next, the part the took me damn near 2 hours to figure out - using CreateRemoteThread
        'We set a pointer to LoadLibrary(LibAddress) in our process, LoadLibrary then puts
        'our Dll(DLLVirtLoc) into the process address.  Easy enough right?
        CreateThread = CreateRemoteThread(ProsH,
                                              vbNull,
                                              0,
                                              LibAddress,
                                              DLLVirtLoc,
                                              0,
                                              ThreadID)
        Verify = 0
    End Function

    Public Function EjectDll(ProcessHandle As IntPtr,
                                 DllHandle As Long)
        Dim LibFreeAddress As Long, CreateEjectThread As Long, EjectThreadId As Long
        'DllHandle = m(ModSrch(DllName)).hModule if u want to go by dll name
        DllHandle = 0
        LibFreeAddress = GetProcAddress(GetModuleHandle("kernel32.dll"),
                                        "FreeLibrary")
        CreateEjectThread = CreateRemoteThread(ProcessHandle,
                                               vbNull,
                                               0,
                                               LibFreeAddress,
                                               DllHandle,
                                               0,
                                               EjectThreadId)
    End Function

#End Region

End Module