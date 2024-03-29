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
'           2021-10-04
'           - New parameter: hc/hideconsole
'           2021-11-18
'           - New parameter /la
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
%VERSION_REVISION = 12

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
   LastAccessTime As Byte
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
'==============================================================================

Function PBMain () As Long
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 10.11.2016
'           - Use own command line parsing instead of buildin PARSE in order
'           to deal with long folder/file names
'           30.01.2017
'           - Resolve absolute and UNC path, if the passed parameter is a relative path
'           and output the information in the application's intro
'           20.06.2018
'           - Expand environment strings included in the path, e.g. %UserProfile%
'           08.11.2019
'           - Use Con.StdOut "..." instead of Print "..." for error messages so that
'           they're also captured if STDOUT is redireted to a file.
'           22.02.2012
'           - New parameter /r (/readonly): force deletion of readonly files
'------------------------------------------------------------------------------
   Local sPath, sTime, sFilePattern, sCmd, sTemp As String
   Local i, j As Dword
   Local lResult, lTemp As Long
   Local vntResult As Variant
   Local udtCfg As ParamsTYPE

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   ' Application intro
   ConHeadline "DeleteFilesOlderThan", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2013-2021", $COMPANY_NAME
   Print ""

   Trace New ".\DeleteFilesOlderThan.tra"

   ' *** Parse the parameters
   ' Initialization and basic checks
   sCmd = Command$

   Local o As IBACmdLine
   Local vnt As Variant

   Let o = Class "cBACmdLine"

   If IsFalse(o.Init(sCmd)) Then
      Print "Couldn't parse parameters: " & sCmd
      Print "Type DeleteFilesOlderThan /? for help"
      Let o = Nothing
      Exit Function
   End If

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0 Then
      ShowHelp
      Exit Function
   End If

   ' Parse the passed parameters
   ' Valid CLI parameters are:
   ' /t= or /time=
   ' /p= or /path=
   ' /f= or /filepattern=
   ' /s or /subfolders
   ' /v or /verbose
   ' /r or /readonly
   ' /fst or /filesmallerthan
   ' /fgt or /filesgreaterthan
   ' /rb or /recyclebin
   ' /pp or /processpriority
   ' /la or /lastaccess
   ' /ddo or /deldirsonly
   ' /dea or /delall

   i = o.ValuesCount

   If (i < 2) Or (i > 8) Then
      Print "Invalid number of parameters."
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Parse CLI parameters

   ' ** Time
   If IsTrue(o.HasParam("t", "time")) Then
      sTemp = Variant$(o.GetValueByName("t", "time"))
      sTime = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** Path
   If IsTrue(o.HasParam("p", "path")) Then
      sTemp = Variant$(o.GetValueByName("p", "path"))
      sPath = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** File pattern
   If IsTrue(o.HasParam("f", "filepattern")) Then
      sTemp = Variant$(o.GetValueByName("f", "filepattern"))
      sFilePattern = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** Recurse subfolders
   If IsTrue(o.HasParam("s", "subfolders")) Then
      vntResult = o.GetValueByName("s", "subfolders")
      'udtCfg.Subfolders = Val(Variant$(vntResult))
      udtCfg.Subfolders = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** Delete to recycle bin
   If IsTrue(o.HasParam("rb", "recyclebin")) Then
      vntResult = o.GetValueByName("rb", "recyclebin")
      'udtCfg.RecycleBin = Sgn(Abs(Val(Variant$(vntResult))))
      udtCfg.RecycleBin = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' Delete readonly files
   If IsTrue(o.HasParam("r", "readonly")) Then
      vntResult = o.GetValueByName("r", "readonly")
      'udtCfg.RecycleBin = Sgn(Abs(Val(Variant$(vntResult))))
      udtCfg.ReadOnly = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' Delete (empty) dirs only
   If IsTrue(o.HasParam("ddo", "deldirsonly")) Then
      vntResult = o.GetValueByName("ddo", "deldirsonly")
      udtCfg.DirsOnly = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' Delete files and directories
   If IsTrue(o.HasParam("dea", "delall")) Then
      vntResult = o.GetValueByName("dea", "delall")
      udtCfg.DirsAndFiles = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' Use LastAccessTime instead of LastWriteTime?
   If IsTrue(o.HasParam("la", "lastaccess")) Then
      vntResult = o.GetValueByName("la", "lastaccess")
      udtCfg.LastAccessTime = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** Set process priority to 'idle' (IDLE_PRIORITY_CLASS) or
   ' 'low' (BELOW_NORMAL_PRIORITY_CLASS)
   If IsTrue(o.HasParam("pp", "processpriority")) Then
      vntResult = o.GetValueByName("pp", "processpriority")
      udtCfg.ProcessPriority = LCase$(Variant$(vntResult))
      If (udtCfg.ProcessPriority <> "i") And (udtCfg.ProcessPriority <> "b") Then
         udtCfg.ProcessPriority = "n"
      End If
   Else
   ' Set default = NORMAL_PRIORITY_CLASS
      udtCfg.ProcessPriority = "n"
   End If

   ' ** Verbose output
   If IsTrue(o.HasParam("v", "verbose")) Then
      vntResult = o.GetValueByName("v", "verbose")
      ' udtCfg.Verbose = Sgn(Abs(Val(Variant$(vntResult))))
      udtCfg.Verbose = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** File size?
   ' Smaller than
   If IsTrue(o.HasParam("fst", "filessmallerthan")) Then
      sTemp = Variant$(o.GetValueByName("fst", "filessmallerthan"))
      udtCfg.FileSize = CalcVal(sTemp)
      udtCfg.CompareFlag = -1
   End If

   ' Greater than
   If IsTrue(o.HasParam("fgt", "filesgreaterthan")) Then
      sTemp = Variant$(o.GetValueByName("fgt", "filesgreaterthan"))
      udtCfg.FileSize = CalcVal(sTemp)
      udtCfg.CompareFlag = 1
   End If

   ' ** Hide console window
   If IsTrue(o.HasParam("hc", "hideconsole")) Then
      vntResult = o.GetValueByName("hc", "hideconsole")
      'udtCfg.RecycleBin = Sgn(Abs(Val(Variant$(vntResult))))
      udtCfg.HideConsole = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If


   ' ** Defaults
   If Len(Trim$(sFilePattern)) < 2 Then
      sFilePattern = "*.*"
   End If

   ' Expand environment variables in path
   sPath = GetEnvironPath(sPath)

   ' Determine if it's a relative or absolute path, i.e. .\MyFolder or C:\MyFolder and/or a UNC share
   Local sPathFull As String
   sPathFull = sPath
   sPathFull = FullPathAndUNC(sPath)

   ' Hide the console window?
   If IsTrue(udtCfg.HideConsole) Then
      Call ShowWindow(Con.Handle, %SW_Hide)
   End If

   ' Echo the CLI parameters
   Con.StdOut "Time               : " & sTime
   Con.StdOut "Folder             : " & sPath;
   ' If path is a relative path, display the full path also
   If LCase$(NormalizePath(sPath)) = LCase$(NormalizePath(sPathFull)) Then
      Con.StdOut ""
   Else
      Con.StdOut " (" & sPathFull & ")"
   End If
   Con.StdOut "File pattern       : " & sFilePattern
   Con.StdOut "Recurse subfolders : " & IIf$(IsTrue(udtCfg.Subfolders), "True", "False")
   Con.StdOut "Verbose            : " & IIf$(IsTrue(udtCfg.Verbose), "True", "False")
   Con.StdOut "Delete readonly    : " & IIf$(IsTrue(udtCfg.ReadOnly), "True", "False")
   Con.StdOut "Delete to Rec. Bin : " & IIf$(IsTrue(udtCfg.RecycleBin), "True", "False")
   ' Con.StdOut "Del. only dirs     : " & IIf$(IsTrue(udtCfg.DirsOnly), "True", "False")
   ' Con.StdOut "Del. files and dirs: " & IIf$(IsTrue(udtCfg.DirsAndFiles), "True", "False")
   Con.StdOut "Base deletion on   : " & IIf$(IsTrue(udtCfg.LastAccessTime), "LastAccessTime", "LastWriteTime")

   Local sPP As String
   sPP = udtCfg.ProcessPriority
   Con.StdOut "Process priority   : " & Switch$(sPP = "i", "Idle", sPP = "b", "Below normal", sPP = "n", "Normal")
   ' File size?
   If udtCfg.CompareFlag <> 0 Then
      Select Case udtCfg.CompareFlag
      Case < 0
         Con.StdOut "Files smaller than : " & sTemp & " (" & Format$(udtCfg.FileSize, "#0,") & " bytes)
      Case > 0
         Con.StdOut "Files greater than : " & sTemp & " (" & Format$(udtCfg.FileSize, "#0,") & " bytes)
      End Select

   End If

   If IsTrue(udtCfg.Verbose) Then
      Call oPTNow.Now()
      Con.StdOut "Current date/time  : " & oPTNow.DateString & ", " & oPTNow.TimeStringFull
   End If

   Con.StdOut ""
   Con.StdOut "Press ";
   Con.Color %INTENSE_WHITE, -1
   Con.StdOut "<ESC>";
   Con.Color %White, -1
   Con.StdOut " at any time to exit."

   StdOut ""

   ' Sanity checks of CLI parameters
   If Len(Trim$(sPath)) < 2 Then
      Con.Color %LITE_RED, -1
      Con.StdOut "Missing folder."
      Con.Color %White, -1
      Con.StdOut ""
      ShowHelp
      Exit Function
   End If

   If Not IsFolder(sPath) Then
      Con.Color %LITE_RED, -1
      Con.StdOut "Folder doesn't exist: " & sPath
      Con.Color %White, -1
      Con.StdOut ""
      ShowHelp
      Exit Function
   End If

   If Len(Trim$(sTime)) < 1 Then
      Con.Color %LITE_RED, -1
      Con.StdOut "Missing time."
      Con.Color %White, -1
      Con.StdOut ""
      ShowHelp
      Exit Function
   End If

   ' Ensure a time unit is given
   If Tally(Right$(LCase$(Trim$(sTime)), 1), Any "dmwy") < 1 Then
      Con.Color %LITE_RED, -1
      Con.StdOut "Missing/invalid time unit: " & Right$(LCase$(Trim$(sTime)), 1) & ". Valid units are d, w, m, y."
      Con.Color %White, -1
      Con.StdOut ""
      ShowHelp
      Exit Function
   End If

   ' Can't use /ddo and /dea together
   If IsTrue(udtCfg.DirsOnly) And IsTrue(udtCfg.DirsAndFiles) Then
      Con.Color %LITE_RED, -1
      Con.StdOut "Parameters /ddo and /dea are mutually exclusive."
      Con.Color %White, -1
      Con.StdOut ""
      ShowHelp
      Exit Function
   End If

   Trace On

   Local qudFileSizeTotal As Quad   ' Total space free by deleted files
   Local qudFolders As Quad         ' # of folders processed
   lResult = DeleteFiles(sPath, sTime, sFilePattern, udtCfg, qudFileSizeTotal, qudFolders)
   Con.StdOut ""
   Con.StdOut "Done."
   Con.StdOut ""
   Con.StdOut "File(s) deleted    : " & Format$(lResult)
   Con.StdOut "Folder(s) processed: " & Format$(qudFolders)

   Local szTemp As AsciiZ * %Max_Path
   Local lRet As Dword

'   Declare Function StrFormatByteSizeA Import "SHLWAPI.DLL" Alias "StrFormatByteSizeA" ( _
'      ByVal dw As Dword _                                  ' __in DWORD dw
'    , ByRef pszBuf As AsciiZ _                             ' __out LPSTR pszBuf
'    , ByVal cchBuf As Dword _                              ' __in UINT cchBuf
'    ) As Dword                                             ' LPSTR


   lRet = StrFormatByteSizeA(CDwd(qudFileSizeTotal), szTemp, Len(szTemp))
   If lRet > 0 Then
      sTemp = Trim$(szTemp, Any Chr$(0,32))
   End If


   ' sTemp = Trim$(GetSizeString(qudFileSizeTotal))
   Con.StdOut "Disk space freed   : " & FormatNumberEx(qudFileSizeTotal, %True) & " bytes" & IIf$(Len(sTemp) > 0, " ~ " & sTemp, "")
   ' Con.StdOut "Disk space freed   : " & FormatNumberEx(qudFileSizeTotal, %True) & " bytes" & IIf$(Len(sTemp) > 0, " ~ " & sTemp, "")

   If IsTrue(udtCfg.Verbose) Then
      Call oPTNow.Now()
      Con.StdOut ""
      Con.StdOut "Current date/time : " & oPTNow.DateString & ", " & oPTNow.TimeStringFull
   End If

   Trace Off
   Trace Close

   Con.StdOut ""

   If IsTrue(glUserExit) Then
      Con.Color %Yellow, -1
      Con.StdOut "Program terminated by user."
      Con.Color %White, -1
   End If

   PBMain = lResult

End Function
'---------------------------------------------------------------------------

Function DeleteFiles(ByVal sPath As String, ByVal sTime As String, ByVal sFilePattern As String, ByVal udtCfg As ParamsTYPE, _
   ByRef qudFileSizeTotal As Quad, ByRef qudFolders As Quad) As Long
'------------------------------------------------------------------------------
'Purpose  : Recursivly scan folders for the file patterns passed and delete files
'           older than sTime
'
'Prereq.  : -
'Parameter: sPath             - Root path for file search
'           sTime             - Time value as passed via parameter. e.g. "2y"
'           sFilePattern      - Files to delete
'           udtCfg            - (Further) configuration
'           qudFileSizeTotal  - (ByRef!) File size in bytes of files that have been deleted
'           qudFolders        - (ByRef!) # of folders processed
'Returns  : The # of deleted files
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 10.11.2016
'           - Use own command line parsing instead of buildin PARSE in order
'           to deal with long folder/file names
'           - 11.11.2016
'           Sum up size of files that were deleted
'           10.03.2017
'           - Allow deletion to recycle bin (/rb)
'           24.08.2018
'           - Display the " - Skipping ..." message only with verbose output enabled (/v)
'           22.02.2012
'           - Enable deletion of readonly files (/r)
'------------------------------------------------------------------------------
   Local sSourceFile, sPattern, sFile, sFileTime As String
   Local sMsg, sTemp, sKeyPress As String
   Local i, lCount As Long
   Local udtDirData As DirData
   Local szSourceFile As WStringZ * %Max_Path
   Local qudFileSize As Quad

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   Local hSearch As Dword                 ' Search handle
   Local udtWFD As WIN32_FIND_DATAW      ' FindFirstFile structure

   Trace On
   Trace Print FuncName$

   ' Set this process' priority
   Select Case LCase$(udtCfg.ProcessPriority)
   Case "i"
      ' %Idle_Priority_Class
      Call SetProcessPriority(%Idle_Priority_Class)
   Case "b"
      Call SetProcessPriority(%BELOW_NORMAL_PRIORITY_CLASS)
   End Select

   For i = 1 To ParseCount(sFilePattern, ";")

      ' Test for program exit
      If IsTrue(CheckUserExit()) Then
         glUserExit = %True
         Exit For
      End If

      Trace Print " -- DeleteFiles sFilePattern: " & Format$(i)

      sMsg = "-- Scanning folder "
      Con.StdOut  sMsg & ShortenPathText(sPath, Con.Screen.Col-(1+Len(sMsg)))

      Trace Print " -- DeleteFiles sPath: " & sPath & " (" & Format$(Len(sPath)) & ")"

      sPattern = Parse$(sFilePattern, ";", i)
      Con.StdOut " - File pattern: " & sPattern

      Trace Print " -- DeleteFiles sPattern: " & sPattern & " (" & Format$(Len(sPattern)) & ")"

      sSourceFile = NormalizePath(sPath) & sPattern
      Trace Print " -- DeleteFiles sSourceFile: " & sSourceFile & " (" & Format$(Len(sSourceFile)) & ")"

      szSourceFile = sSourceFile

      hSearch = FindFirstFileW(szSourceFile, udtWFD)  ' Get search handle, if success
      If hSearch <> %INVALID_HANDLE_VALUE Then        ' Loop through directory for files

         ' At least the starting folder exists
         Incr qudFolders

         Do

            qudFileSize = 0

            If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY Then ' If not directory bit is set (files only here...)

               sFile = Remove$(udtWFD.cFileName, Any Chr$(0))

               If IsTrue(udtCfg.Verbose) Then
                  sFileTime = GetFileTimeStr(udtWFD)
               End If

               If IsTrue(IsDeleteMatch(sTime, udtWFD, udtCfg)) Then

                  sMsg = "  - Deleting "
                  Con.StdOut  sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                  If IsTrue(udtCfg.Verbose) Then
                     Con.StdOut "    Time stamp: " & sFileTime;
                  End If
                  Incr lCount

                  Try
                     ' Get the file size before deleting it
                     qudFileSize = GetThisFileSize(udtWFD)
                     If IsFalse(udtCfg.RecycleBin) Then
                        Kill NormalizePath(sPath) & sFile
                     Else
                        Call Delete2RecycleBin(NormalizePath(sPath) & sFile)
                     End If
                     If IsTrue(udtCfg.Verbose) Then
                        Con.StdOut " - File size: " & FormatNumberEx(qudFileSize, %True) & " bytes"
                     End If

                  Catch
                     ' Delete readonly files, if this caused the error ("Permission denied")

                     Trace Print " -- Err          : " & Format$(Err)
                     Trace Print " -- Cfg.ReadOnly : " & Format$(udtCfg.ReadOnly)
                     Trace Print " -- File attr.   : " & Format$(udtWFD.dwFileAttributes)
                     Trace Print " -- File attr AND: " & Format$(udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_READONLY)
                     Trace Print " -- File attr  OR: " & Format$(udtWFD.dwFileAttributes Or %FILE_ATTRIBUTE_READONLY)

                     If (Err = 70) And (IsTrue(udtCfg.ReadOnly)) And (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_READONLY) Then
                        SetAttr NormalizePath(sPath) & sFile, (udtWFD.dwFileAttributes - %FILE_ATTRIBUTE_READONLY)
                        ErrClear
                        Try
                           If IsFalse(udtCfg.RecycleBin) Then
                              Kill NormalizePath(sPath) & sFile
                           Else
                              Call Delete2RecycleBin(NormalizePath(sPath) & sFile)
                           End If
                        Catch
                           Con.Color 12, -1
                           sMsg = "  - ERROR: can't delete "
                           Con.StdOut sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg))) & " - " & Error$(Err)
                           If IsTrue(udtCfg.Verbose) Then
                              Con.StdOut ""
                           End If
                           Con.Color 7, -1
                           Decr lCount
                        End Try

                     Else

                        Con.Color 12, -1
                        sMsg = "  - ERROR: can't delete "
                        Con.StdOut sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg))) & " - " & Error$(Err)
                        If IsTrue(udtCfg.Verbose) Then
                           Con.StdOut ""
                        End If
                        Con.Color 7, -1
                        Decr lCount
                     End If

                  End Try

               Else

                  If IsTrue(udtCfg.Verbose) Then
                     sMsg = "  - Skipping "
                     Con.StdOut sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                     If IsTrue(udtCfg.Verbose) Then
                        Con.StdOut "    Time stamp: " & sFileTime
                     End If
                  End If   '// If IsTrue(udtCfg.Verbose)

               End If

            End If   '// If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY

            ' Sum up file size
            qudFileSizeTotal = qudFileSizeTotal + qudFileSize

            ' Test for program exit
            If IsTrue(CheckUserExit()) Then
               glUserExit = %True
               Exit Do
            End If

         Loop While FindNextFileW(hSearch, udtWFD)

         FindClose hSearch

      End If   '// If hSearch <> %INVALID_HANDLE_VALUE

      Con.StdOut ""


      If IsTrue(udtCfg.Subfolders) Then  ' if to search in subdirectories.

         szSourceFile = NormalizePath(sPath) & "*"
         hSearch = FindFirstFileW(szSourceFile, udtWFD)

         If hSearch <> %INVALID_HANDLE_VALUE Then

            Do

               If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) = %FILE_ATTRIBUTE_DIRECTORY _
                  And (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_HIDDEN) = 0 Then  ' If dirs, but not hidden..

                  If udtWFD.cFileName <> "." And udtWFD.cFileName <> ".." Then          ' Not these..
                     lCount = lCount + DeleteFiles(NormalizePath(sPath) & RTrim$(udtWFD.cFileName, $Nul), sTime, sFilePattern, udtCfg, qudFileSizeTotal, qudFolders)
                  End If

               End If

               ' Test for program exit
               If IsTrue(CheckUserExit()) Then
                  glUserExit = %True
                  Exit Do
               End If

            Loop While FindNextFileW(hSearch, udtWFD)

            FindClose hSearch

         End If

      End If

   Next i

   DeleteFiles = lCount

End Function
'---------------------------------------------------------------------------

Function IsDeleteMatch(ByVal sTime As String, ByVal udt As DirData, ByVal udtCfg As ParamsTYPE) As Long
'------------------------------------------------------------------------------
'Purpose  : Determine if a file matches the deletion criterias
'
'Prereq.  : -
'Parameter: sTime    - Time value as passed via parameter
'           udt      - File information about the current file (Win32_Find_Data)
'           udtCfg   - Parameters passed
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 10.11.2016
'           - Use own command line parsing instead of buildin PARSE in order
'           to deal with long folder/file names
'           11.11.2016
'           - Sum up size of files that were deleted
'           30.01.2017
'           - Compare file size in addition to file time
'           2021-11-18
'           - Calculated time by LastAccessTime
'------------------------------------------------------------------------------
   Local oPTFile, oPTNow As IPowerTime
   Local dwValue As Dword
   Local sUnit As String
   Local unFS As FileSizeUNION

   Let oPTFile = Class "PowerTime":Let oPTNow = Class "PowerTime"
   Call oPTNow.Now()

   If IsTrue(udtCfg.LastAccessTime) Then
      oPTFile.FileTime = udt.LastAccessTime
   Else
      oPTFile.FileTime = udt.LastWriteTime
   End If

   ' Extract value and unit of the time parameter
   dwValue = CDwd(Val(sTime))
   Select Case LCase$(Right$(sTime, 1))
   Case "d"
      ' Days
      Call oPTNow.AddDays(-dwValue)
   Case "m"
      ' Month
      Call oPTNow.AddMonths(-dwValue)
   Case "w"
      ' Weeks
      Call oPTNow.AddDays(-dwValue * 7)
   Case "y"
      Call oPTNow.AddYears(-dwValue)
   End Select

   unFS.Part.Lo = udt.FileSizeLow
   unFS.Part.Hi = udt.FileSizeHigh

   Trace Print "FileTime: " & Format$(oPTFile.FileTime)
   Trace Print "NowTime : " & Format$(oPTNow.FileTime)
   Trace Print " - udtCfg.FileSize: " & Format$(udtCfg.FileSize)
   Trace Print " - unFS.Full      : " & Format$(unFS.Full)

   ' Assume false
   IsDeleteMatch = %False

   If oPTFile.FileTime <= oPTNow.FileTime Then
      If udtCfg.CompareFlag = 0 Then
         IsDeleteMatch = %True
      Else
         If ((udtCfg.CompareFlag < 0) And (unFS.Full < udtCfg.FileSize)) Or ((udtCfg.CompareFlag > 0) And (unFS.Full > udtCfg.FileSize)) Then
            IsDeleteMatch = %True
            Trace Print "  - Size: True"
         End If
      End If
   End If

   ' *** Debug
   'IsDeleteMatch = %False

End Function
'---------------------------------------------------------------------------

Function GetFileTimeStr(ByVal udt As DirData) As String
'------------------------------------------------------------------------------
'Purpose  : Formats a given FILETIME structure's value as a readable sting
'
'Prereq.  : -
'Parameter: udt   - File information about the current file (Win32_Find_Data)
'Returns  : Localized date/time string
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local oPTFile As IPowerTime

   Let oPTFile = Class "PowerTime"
   oPTFile.FileTime = udt.LastWriteTime

   GetFileTimeStr = oPTFile.DateString & ", " & oPTFile.TimeStringFull

End Function
'---------------------------------------------------------------------------

Sub ShowHelp
'------------------------------------------------------------------------------
'Purpose  : Show usage help
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Con.StdOut ""
   Con.StdOut "DeleteFilesOlderThan"
   Con.StdOut "--------------------"
   Con.StdOut "DeleteFilesOlderThan deletes files matching the passed file pattern and which are older than the given time specification from a folder."
   Con.StdOut ""
   Con.StdOut "Usage:   DeleteFilesOlderThan _"
   Con.StdOut "            /time=<time specification> /path=<folder to delete files from> [/filepattern=<files to delete>[;<files to delete>]] _"
   Con.StdOut "            [/subfolders=0|1] [/filessmallerthan=|/filesgreaterthan=<file size>] [/recyclebin=0|1] [/processpriority=i|b]"
   Con.StdOut "            [/deldirsonly|/delall]"
   Con.StdOut "  or     DeleteFilesOlderThan /t=<time specification> /p=<folder to delete files from> [/f=<files to delete>[;<files to delete>]] [/s=0|1] [/fst=|/fgt=<file size>] [/rb=0|1]  [/pp=i|b]"
   Con.StdOut "e.g.     DeleteFilesOlderThan /time=2d /path=D:\MyTarget"
   Con.StdOut "         DeleteFilesOlderThan /t=3w /p=C:\MyTarget\Data /f=*.txt /s=1"
   Con.StdOut ""
   Con.StdOut "Pressing <ESC> any time will exit the program."
   Con.StdOut ""
   Con.StdOut "Parameters"
   Con.StdOut "----------"
   Con.StdOut "/t or   /time             = time specification"
   Con.StdOut "/p or   /path             = (start) folder"
   Con.StdOut "                            If path includes spaces or other special characters, enclose it in double quotation marks (" & Chr$(34) &")"
   Con.StdOut "                            e.g. /p=" & Chr$(34) & "C:\My Data Folder\Old Backups" & Chr$(34)
   Con.StdOut "/f or   /filepattern      = file pattern"
   Con.StdOut "         If omitted, all files are scanned (equals /f=*.*)."
   Con.StdOut "/s or   /subfolders       = recurse subfolders yes(1) or no (0)"
   Con.StdOut "         If omitted, only the folder passed via /p is scanned for matching files (equals /s=0)."
   Con.StdOut "/rb or  /recyclebin       = delete to recycle bin instead of permanently delete."
   Con.StdOut "         If omitted, defaults to 0 = delete files permanently."
   Con.StdOut "/pp or  /processpriority  = Lower this process' priority in order to consume less (mainly CPU) resources."
   Con.StdOut "         Valid values are i = Idle (lowest possible priority) or b = Below Normal.
   Con.StdOut "/hc or  /hideconsole      = Hide the application's (console) window? Yes(1) or no(0). Defaults to no."
   Con.StdOut "/la or  /lastaccess       = Compare to a file's LastAccessTime instead of its LastWriteTime."
   Con.StdOut "/fst or /filessmallerthan = only delete files smaller than the specified file size (see below how to pass file sizes)."
   Con.StdOut "/fgt or /filesgreaterthan = only delete files greater than the specified file size (see below how to pass file sizes)."
   Con.StdOut ""
   Con.StdOut "Please note that you may only use *either* /fst or /fgt. You can't use both parameters. If you happen to pass both parameters, the last one 'wins'."
   Con.StdOut ""
'   Con.StdOut "/ddo or /deldirsonly      = delete (empty) directories only."
'   Con.StdOut "/dea or /delall           = delete directories and files."
'   Con.StdOut ""
'   Con.StdOut "Both options (/ddo, /dea) utilize the provided file pattern (/f). Suppose the file pattern being /f=tmp*.tm?, then:"
'   Con.StdOut "/ddo won't delete ANY files at all. And it will only delete empty directories matching the pattern."
'   Con.StdOut "/dea will delete all files matching the pattern. And all directories matching it, with *everything* in it. The file pattern doesn't apply to those.
'   Con.StdOut "/ddo and /dea can't be used together."
'   Con.StdOut ""
   Con.StdOut "You may specify more than one file pattern for the parameter /f by using ; (semicolon) as a separator, e.g."
   Con.StdOut "       /f=*.doc;*.rtf -> deletes all *.doc and all *.rtf files from the specified folder."
   Con.StdOut "       /f=Backup*.bak;Log*.trn -> deletes all Backup*.bak and all Log*.trn files from the specified folder."
   Con.StdOut ""
   Con.StdOut "Allowed time specification units for parameter /t are:
   Con.StdOut "        d = day   e.g. 1d"
   Con.StdOut "        w = week  e.g. 2w"
   Con.StdOut "        m = month e.g. 3m"
   Con.StdOut "        y = year  e.g. 4y"
   Con.StdOut ""
   Con.StdOut "Allowed file size units:"
   Con.StdOut "   <empty> = Byte, e.g. 100"
   Con.StdOut "   kb = Kilobyte, e.g. 100kb"
   Con.StdOut "   mb = Megabyte, e.g. 100mb"
   Con.StdOut "   gb = Gigabyte, e.g. 100gb"
   Con.StdOut "   tb = Terabyte, e.g. 100tb"
   Con.StdOut ""
   Con.StdOut "Please note: 1 KB = 1024 byte, 1 MB = 1024 KB etc."
   Con.StdOut ""

End Sub
'---------------------------------------------------------------------------

Function GetThisFileSize(ByVal udtFileSize As WIN32_FIND_DATAW) As Quad
'------------------------------------------------------------------------------
'Purpose  : Derive a file's size from WIN32_FIND_DATAW
'
'Prereq.  : -
'Parameter: -
'Returns  : File size in bytes
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Function = udtFileSize.nFileSizeHigh * &H0100000000 + udtFileSize.nFileSizeLow
End Function
'---------------------------------------------------------------------------

Function GetSizeString(ByVal q As Quad) As String
'------------------------------------------------------------------------------
'Purpose  : Format a file's size as a string, using common unit abbreviations,
'           e.g. 1TB, 2GB
'
'Prereq.  : -
'Parameter: -
'Returns  : File size in bytes
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sSize As String
   Local qudDivisor As Quad

   Trace On
   Trace Print "q: " & Format$(q)

   Do While q > 0

      If (q \ 1024&&^4) > 0 Then
      ' TB
         qudDivisor = q \ 1024&&^4
         q = q - (qudDivisor * 1024&&^4)
         Trace Print "TB: " & Format$(qudDivisor)
         sSize = sSize & Format$(qudDivisor) & "TB "
      ElseIf  q \ 1024&&^3 > 0 Then
      ' GB
         qudDivisor = q \ 1024&&^3
         q = q - (qudDivisor * 1024&&^3)
         Trace Print "GB: " & Format$(qudDivisor)
         sSize = sSize & Format$(qudDivisor) & "GB "
      ElseIf  q \ 1024&&^2 > 0 Then
      ' MB
         qudDivisor = q \ 1024&&^2
         q = q - (qudDivisor * 1024&&^2)
         Trace Print "MB: " & Format$(qudDivisor)
         sSize = sSize & Format$(qudDivisor) & "MB "
      ElseIf  q \ 1024&&^1 > 0 Then
      ' KB
         qudDivisor = q \ 1024&&^1
         q = q - (qudDivisor * 1024&&^1)
         Trace Print "KB: " & Format$(qudDivisor)
         sSize = sSize & Format$(qudDivisor) & "KB "
      Else
      ' B
         qudDivisor = q \ 1024&&^0
         q = q - (qudDivisor * 1024&&^0)
         Trace Print "B: " & Format$(qudDivisor)
         sSize = sSize & Format$(qudDivisor) & "B"
      End If

   Loop

   Function = Trim$(sSize)

End Function
'---------------------------------------------------------------------------

Function CalcVal (ByVal sValue As String) As Quad
'------------------------------------------------------------------------------
'Purpose  : Calculate the value in bytes of a size expression
'           e.g. 1kb -> 1024
'
'Prereq.  : -
'Parameter: Size value string
'Returns  : Equiv. Size in bytes
'Note     : Uses proper power of two for calculation instead of industry
'           standard 1000
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   sValue = LCase$(sValue)
   Select Case Right$(sValue, 2)
   Case "kb"
      CalcVal = Val(sValue) * 1024&&
   Case "mb"
      CalcVal = Val(sValue) * 1024&&^2
   Case "gb"
      CalcVal = Val(sValue) * 1024&&^3
   Case "tb"
      CalcVal = Val(sValue) * 1024&&^4
   Case Else
      CalcVal = Val(sValue)
   End Select

End Function
'---------------------------------------------------------------------------

Function CheckUserExit(Optional ByVal vntKeys As Variant) As Long
'------------------------------------------------------------------------------
'Purpose  : Check if user pressed ESC and signals program tgermination
'
'Prereq.  : -
'Parameter: lExitFlag   - (ByRef!) Will be set to %True, if program termination was detected.
'           vntKeys     - One or more keys that should be checked in addition to ESC,
'                       e.g. "cx" will check both for lower- and uppercase "c" and
'                       "x" and if detected, signal program termination
'Returns  : True / False   - ESC was pressed?
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sPressed, sTemp As String
   Local i As Long

   Static sKeys As String
   Static lBeenHere As Long

   ' Build the string only once

   If IsFalse(lBeenHere) Then

      ' ESC always exits
      sKeys = $Esc

      If IsFalse(IsMissing(vntKeys)) Then

         ' Add lower and uppercase variant to the string
         sTemp = Variant$(vntKeys)
         For i = 1 To Len(sTemp)
            sKeys &= sKeys & LCase$(Mid$(sTemp, i, 1)) & UCase$(Mid$(sTemp, i, 1))
         Next i

      End If

      lBeenHere = %True

   End If

   sPressed = Con.InKey$

   ' Any of the termination keys pressed?
   CheckUserExit = Tally(sPressed, Any sKeys)

End Function
'------------------------------------------------------------------------------
