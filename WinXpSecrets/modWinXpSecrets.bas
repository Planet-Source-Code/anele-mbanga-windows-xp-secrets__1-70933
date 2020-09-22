Attribute VB_Name = "modWinXpSecrets"
Option Explicit

Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibfName As String) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal xC As Long)
Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long

Private Enum RegisterServerEnum
    DllUnregisterServer = 0
    DllRegisterServer = 1
End Enum

Sub Main()
    On Error Resume Next
    Dim regFile As String
    regFile = GetSysDir & "\Regtool5.dll"
    If File_Exists(regFile) = False Then
        ResourceFile_SaveItemToDisk 101, "CUSTOM", regFile
        DoEvents
        regFile = RegisterServer(regFile, DllRegisterServer)
    End If
    frmWinXp.Show
    Err.Clear
End Sub

Private Function GetSysDir() As String
    On Error Resume Next
    Dim lpBuffer As String * 255
    Dim Length As Long
    Length = apiGetSystemDirectory(lpBuffer, Len(lpBuffer))
    GetSysDir = Left$(lpBuffer, Length)
    Err.Clear
End Function

Private Function File_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    File_Exists = fs.FileExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function

Private Function ResourceFile_SaveItemToDisk(ByVal iResourceNum As Integer, ByVal sResourceType As String, ByVal sDestFileName As String) As Long
    On Error Resume Next
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    On Error GoTo SaveResItemToDisk_err
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    iFileNumOut = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNumOut
        Put #iFileNumOut, , bytResourceData
    Close #iFileNumOut
    ResourceFile_SaveItemToDisk = 0
    Err.Clear
    Exit Function
SaveResItemToDisk_err:
    ResourceFile_SaveItemToDisk = Err.Number
    Err.Clear
End Function

Private Function RegisterServer(fName As String, RegFunc As RegisterServerEnum) As String
    On Error Resume Next
    Dim regLib As Long
    Dim process As Long
    Dim succeed As Long
    Dim h1 As Long
    Dim xC As Long
    Dim ID As Long
    Dim P As String
    Select Case RegFunc
    Case DllUnregisterServer
        P = "DllUnregisterServer"
    Case DllRegisterServer
        P = "DllRegisterServer"
    End Select
    regLib = LoadLibraryRegister(fName)
    If regLib = 0 Then
        RegisterServer = "Error"
        Err.Clear
        Exit Function
    End If
    process = GetProcAddressRegister(regLib, P)
    If process = 0 Then
        RegisterServer = "Error"
    Else
        h1 = CreateThreadForRegister(ByVal 0&, 0&, ByVal process, ByVal 0&, 0&, ID)
        If h1 = 0 Then
            RegisterServer = "Error"
        Else
            succeed = (WaitForSingleObject(h1, 10000) = 0)
            If succeed Then
                CloseHandle h1
                RegisterServer = "Ok"
            Else
                GetExitCodeThread h1, xC
                ExitThread xC
                RegisterServer = "Error"
            End If
        End If
    End If
    FreeLibraryRegister regLib
    Err.Clear
End Function

Public Function Registry_Save(KeyRoot As REGTool5.REGToolRootTypes, KeyName As String, ValueName As String, ValueData As String) As Boolean
    On Error Resume Next
    ' saves a key to the registry
    ' Example: SaveReg HKEY_CURRENT_USER, "Software\Kimmo\ODBC Drivers List", "Driver", "SQL Server"
    Registry_Save = REGTool5.UpdateKey(KeyRoot, KeyName, ValueName, ValueData)
    Err.Clear
End Function

Public Function Registry_Read(KeyRoot As REGTool5.REGToolRootTypes, KeyName As String, ValueName As String, ByRef ValueData As String) As Boolean
    On Error Resume Next
    ' reads a key from the registry
    ' Example: ReadReg HKEY_CURRENT_USER, "Software\Kimmo\ODBC Drivers List", "Driver", "SQL Server"
    Registry_Read = REGTool5.GetKeyValue(KeyRoot, KeyName, ValueName, ValueData)
    Err.Clear
End Function

Public Function Registry_Delete(KeyRoot As REGTool5.REGToolRootTypes, KeyName As String) As Boolean
    On Error Resume Next
    ' reads a key from the registry
    Registry_Delete = REGTool5.DeleteKey(KeyRoot, KeyName)
    Err.Clear
End Function

