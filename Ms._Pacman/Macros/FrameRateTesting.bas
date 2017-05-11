Attribute VB_Name = "FrameRateTesting"
Option Explicit

Public Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Public Const sCPURegKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public StartMT As Double
Public EndMT As Double

Function MicroTimer() As Double
    '
    'returns seconds
    '
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency 'get ticks/sec
    getTickCount cyTicks1 'get ticks
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency 'calc seconds

End Function
