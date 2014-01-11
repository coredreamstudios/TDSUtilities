@echo off
:: This will copy the BMP adn DLL files to the
:: Windows\System directory and then register
:: the DLL to the system

::copy vbcmts.bmp c:\winnt\system /y
::copy vbcmts.dll c:\winnt\system /y
regsvr32.exe "c:\winnt\system\vbcmts.dll"

::C:\Documents and Settings\crock9l\My Documents\Visual Studio Projects\VBComment\Dll