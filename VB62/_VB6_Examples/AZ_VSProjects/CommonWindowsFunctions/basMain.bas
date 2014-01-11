Attribute VB_Name = "Module1"
Option Explicit

Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Declare Function GetEnvironmentVariable Lib "KERNEL32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, _
                                                                                                                             ByVal lpBuffer As String, _
                                                                                                                             ByVal nSize As Long) As Long


Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal lpvDest As String, ByVal lpvSource As String, cbLen As Long)

Declare Function GetEnvironmentStrings Lib "KERNEL32" Alias "GetEnvironmentStringsA" () As Long

Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long

Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                                                                                              nSize As Long) As Long

Declare Function GetTickCount Lib "KERNEL32" () As Long
