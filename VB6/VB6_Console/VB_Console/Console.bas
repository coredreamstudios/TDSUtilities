Attribute VB_Name = "basConsole"


Option Explicit
Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long

Public Const STD_OUTPUT_HANDLE = -11&
Public hConsole As Long
Public Function ConsoleHandler(ByVal CtrlType As Long) As Long
    
    ConsoleHandler = 1  'Keep track of the console, it'll survive even Ctrl+Break
    
End Function
