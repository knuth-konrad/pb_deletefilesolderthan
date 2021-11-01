'------------------------------------------------------------------------------
'Purpose  : Delete files older than a given date/time
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 2013
'   Source: -
'  Changed: 15.11.2016
'           - Provide information about the amount of disk space freed.
'           30.01.2017
'           - New optional parameters: /fst, /filessmallerthan and /fgt /filesgreaterthan
'           - Resolve absolute and UNC path, if the passed parameter is a relative path
'           and output the information in the application's intro
'           10.03.2017
'           - Allow deletion to recycle bin (/rb)
'           04.05.2017
'           - #Break On to prevent console context menu changes
'           10.05.2017
'           - Switch from source code to SLL include
'           15.05.2017
'           - Application manifest added
'           10.07.2017
'           - Recompile because of lib changes
'           18.06.2018
'           - New parameter: /pp=i|b (ProcessPriority=Idle or Below normal)
'           20.06.2018
'           - Exit the current run with <ESC>
'           04.12.2018
'           - Format numbers with proper locale
'           10-04-2021
'           - New parameter: hc/hideconsole
'------------------------------------------------------------------------------
#Compile Exe ".\DeleteFilesOlderThan.exe"
#Option Version5
#Break On
#Dim All

#Link "baCmdLine.sll"

#Debug Error Off
#Tools Off

%VERSION_MAJOR = 1
%VERSION_MINOR = 8
%VERSION_REVISION = 10

' Version Resource information
#Include ".\DeleteFilesOlderThanRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
' Console colors
%Green = 2
%Red = 4
%White = 7
%Yellow = 14
%LITE_GREEN = 10
%LITE_RED = 12
%INTENSE_WHITE = 15
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Type ParamsTYPE
   Subfolders As Byte
   Verbose As Byte
   CompareFlag As Byte
   FileSize As Quad
   RecycleBin As Byte
   ProcessPriority As String * 1
   DirsOnly As Byte
   DirsAndFiles As Byte
   ReadOnly As Byte
   HideConsole As Byte
End Type

Type FileSizeTYPE
   Lo As Dword
   Hi As Dword
End Type

Union FileSizeUNION
   Full As Quad
   Part As FileSizeTYPE
End Union
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "win32api.inc"
#Include "sautilcc.inc"       ' General console helpers

Declare Function PathIsDirectoryEmptyA Import "SHLWAPI.DLL" Alias "PathIsDirectoryEmptyA" ( _
   ByRef pszPath As AsciiZ _                            ' __in LPCSTR pszPath
   ) As Long                                            ' BOOL

Declare Function StrFormatByteSizeA Import "SHLWAPI.DLL" Alias "StrFormatByteSizeA" ( _
   ByVal dw As Dword _                                  ' __in DWORD dw
 , ByRef pszBuf As AsciiZ _                             ' __out LPSTR pszBuf
 , ByVal cchBuf As Dword _                              ' __in UINT cchBuf
 ) As Dword                                             ' LPSTR

'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
' User signaled program exit
Global glUserExit As Long
