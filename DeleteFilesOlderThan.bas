#Compile Exe "DeleteFilesOlderThan.exe"
#Option Version5
#Dim All

#Debug Error Off
#Tools Off

%VERSION_MAJOR = 1
%VERSION_MINOR = 4
%VERSION_REVISION = 1

' Version information resource
#Include ".\DeleteFilesOlderThanRes.inc"

#Include Once "win32api.inc"
#Include "sautilcc.inc"
'---------------------------------------------------------------------------

Function PBMain () As Long

   Local sPath, sTime, sFilePattern, sTemp As String
   Local i, j As Dword
   Local lSubfolders, lResult, lVerbose As Long

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   ConHeadline "DeleteFilesOlderThan", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2013, 2016", $COMPANY_NAME
   Print ""

   Trace New ".\DeleteFilesOlderThan.tra"

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0 Then
      ShowHelp
      Exit Function
   End If

   ' Valid CLI parameters are:
   ' /t= or /time=
   ' /p= or /path=
   ' /f= or /filepattern=
   ' /s= or /subfolders=
   ' /v= or /verbose
   i = ParseCount(Command$, "/") - 1

   If (i < 2) Or (i > 5) Then
      Print "Invalid number of parameters."
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Parse CLI parameters
   For j = 1 To i

      If InStr(LCase$(Command$(j)),"/t=") > 0 Or InStr(LCase$(Command$(j)),"/time=") > 0 Then
         sTemp = LCase$(Command$(j))
         Replace "/time" With "/t" In sTemp
         sTime = Remove$(Remain$(sTemp, "/t="), $Dq)

      ElseIf InStr(LCase$(Command$(j)),"/p=") > 0 Or InStr(LCase$(Command$(j)),"/path=") > 0 Then
         sTemp = LCase$(Command$(j))
         Replace "/path" With "/p" In sTemp
         sPath = Remove$(Remain$(sTemp, "/p="), $Dq)

      ElseIf InStr(LCase$(Command$(j)),"/f=") > 0 Or InStr(LCase$(Command$(j)),"/filepattern=") > 0 Then
         sTemp = LCase$(Command$(j))
         Replace "/filepattern" With "/f" In sTemp
         sFilePattern = Remove$(Remain$(sTemp, "/f="), $Dq)

      ElseIf InStr(LCase$(Command$(j)),"/s=") > 0 Or InStr(LCase$(Command$(j)),"/subfolders=") > 0 Then
         sTemp = LCase$(Command$(j))
         Replace "/subfolders" With "/s" In sTemp
         lSubfolders = Abs(Val(Remove$(Remain$(sTemp, "/s="), $Dq)))

      ElseIf InStr(LCase$(Command$(j)),"/v=") > 0 Or InStr(LCase$(Command$(j)),"/verbose=") > 0 Then
         sTemp = LCase$(Command$(j))
         Replace "/verbose" With "/v" In sTemp
         lVerbose = Abs(Val(Remove$(Remain$(sTemp, "/v="), $Dq)))
      End If

   Next j

   If Len(Trim$(sFilePattern)) < 2 Then
      sFilePattern = "*.*"
   End If

   ' Echo the CLI parameters
   Con.StdOut "Time              : " & sTime
   Con.StdOut "Folder            : " & sPath
   Con.StdOut "File pattern      : " & sFilePattern
   Con.StdOut "Recurse subfolders: " & Choose$(lSubfolders, "True" Else "False")
   Con.StdOut "Verbose           : " & Choose$(lVerbose, "True" Else "False")

   If IsTrue(lVerbose) Then
      Call oPTNow.Now()
      Con.StdOut "Current date/time : " & oPTNow.DateString & ", " & oPTNow.TimeStringFull
   End If

   StdOut ""

   ' Sanity checks of CLI parameters
   If Len(Trim$(sPath)) < 2 Then
      Print "Missing folder."
      Print ""
      ShowHelp
      Exit Function
   End If

   If Not IsFolder(sPath) Then
      Print "Folder doesn't exist:" & sPath
      Print ""
      ShowHelp
      Exit Function
   End If

   If Len(Trim$(sTime)) < 1 Then
      Print "Missing time."
      Print ""
      ShowHelp
      Exit Function
   End If

   If Len(Trim$(sTime)) > 1 Then
      sTime = Trim$(sTime)
      If Verify(Right$(sTime, 1), "dmwy") > 0 Then
         Print "Invalid time unit."
         Print ""
         ShowHelp
         Exit Function
      End If
   End If

   Trace On
   lResult = DeleteFiles(sPath, sTime, sFilePattern, lSubfolders, lVerbose)
   StdOut ""
   StdOut "Done. " & Format$(lResult) & " file(s) deleted."
   Trace Off

   Trace Close

   StdOut ""

   PBMain = lResult

End Function
'---------------------------------------------------------------------------

Function DeleteFiles(ByVal sPath As String, ByVal sTime As String, ByVal sFilePattern As String, ByVal lSubfolders As Long, ByVal lVerbose As Long) As Long

   Local sSourceFile, sPattern, sFile, sFileTime As String
   Local sMsg, sTemp As String
   Local i, lCount As Long
   Local udtDirData As DirData
   Local szSourceFile As WStringZ * %Max_Path

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   Local hSearch As Dword                ' Search handle
   Local udtWFD As WIN32_FIND_DATAW      ' FindFirstFile structure


   For i = 1 To ParseCount(sFilePattern, ";")

      sMsg = "-- Scanning folder "
      Con.StdOut  sMsg & ShortenPathText(sPath, Con.Screen.Col-(1+Len(sMsg)))

      sPattern = Parse$(sFilePattern, ";", i)
      Con.StdOut " - File pattern: " & sPattern

      sSourceFile = NormalizePath(sPath) & sPattern
      szSourceFile = sSourceFile

      hSearch = FindFirstFileW(szSourceFile, udtWFD)  ' Get search handle, if success
      If hSearch <> %INVALID_HANDLE_VALUE Then        ' Loop through directory for files

         Do

            If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY Then ' If not directory bit is set (files only here...)

               sFile = Remove$(udtWFD.cFileName, Any Chr$(0))

               If IsTrue(lVerbose) Then
                  sFileTime = GetFileTimeStr(udtWFD)
               End If

               ' Test the file pattern against the file name retrieved.
               ' After a positive match, check the file's time stamp and delete it,
               ' if it falls within the given time frame.
               If IsTrue(IsDeleteMatch(sTime, udtWFD)) Then

                  sMsg = "  - Deleting "
                  Con.StdOut  sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                  If IsTrue(lVerbose) Then
                     Con.StdOut "    Time stamp: " & sFileTime
                  End If
                  Incr lCount

                  Try
                     Kill NormalizePath(sPath) & sFile
                  Catch
                     Con.Color 12, -1
                     sMsg = "  - ERROR: can't delete "
                     Con.StdOut  sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                     If IsTrue(lVerbose) Then
                        Con.StdOut "    Time stamp: " & sFileTime
                     End If
                     Con.Color 7, -1
                     Decr lCount
                  End Try

               Else

                  sMsg = "  - Skipping "
                  Con.StdOut  sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                  If IsTrue(lVerbose) Then
                     Con.StdOut "    Time stamp: " & sFileTime
                  End If

               End If

            End If   '// If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY

         Loop While FindNextFileW(hSearch, udtWFD)

         FindClose hSearch

      End If   '// If hSearch <> %INVALID_HANDLE_VALUE

      Con.StdOut ""


      If IsTrue(lSubfolders) Then  ' If to recurse subdirectories.

         szSourceFile = NormalizePath(sPath) & "*"
         hSearch = FindFirstFileW(szSourceFile, udtWFD)

         If hSearch <> %INVALID_HANDLE_VALUE Then

            Do

               If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) = %FILE_ATTRIBUTE_DIRECTORY _
                  And (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_HIDDEN) = 0 Then  ' If dirs, but not hidden..

                  If udtWFD.cFileName <> "." And udtWFD.cFileName <> ".." Then          ' Not these..
                     lCount = lCount + DeleteFiles(NormalizePath(sPath) & RTrim$(udtWFD.cFileName, $Nul), sTime, sFilePattern, lSubfolders, lVerbose)
                  End If

               End If

            Loop While FindNextFileW(hSearch, udtWFD)

            FindClose hSearch

         End If

      End If

   Next i

   DeleteFiles = lCount

End Function
'---------------------------------------------------------------------------

Function IsDeleteMatch(ByVal sTime As String, ByVal udt As DirData) As Long

   Local oPTFile, oPTNow As IPowerTime
   Local dwValue As Dword
   Local sUnit As String

   Let oPTFile = Class "PowerTime":Let oPTNow = Class "PowerTime"
   Call oPTNow.Now()
   oPTFile.FileTime = udt.LastWriteTime

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

   Trace Print "FileTime: " & Format$(oPTFile.FileTime)
   Trace Print "NowTime : " & Format$(oPTNow.FileTime)

   If oPTFile.FileTime <= oPTNow.FileTime Then
      IsDeleteMatch = %True
   Else
      IsDeleteMatch = %False
   End If

End Function
'---------------------------------------------------------------------------

Function GetFileTimeStr(ByVal udt As DirData) As String

   Local oPTFile As IPowerTime

   Let oPTFile = Class "PowerTime"
   oPTFile.FileTime = udt.LastWriteTime

   GetFileTimeStr = oPTFile.DateString & ", " & oPTFile.TimeStringFull

End Function
'---------------------------------------------------------------------------

Sub ShowHelp

   Con.StdOut ""
   Con.StdOut "DeleteFilesOlderThan"
   Con.StdOut "--------------------"
   Con.StdOut "DeleteFilesOlderThan deletes files matching the passed file pattern and which are older than the given time specification from a folder."
   Con.StdOut ""
   Con.StdOut "Usage:   DeleteFilesOlderThan /time=<time specification> /path=<folder to delete files from> [/filepattern=<files to delete>[;<files to delete>]] [/subfolders=0|1]"
   Con.StdOut "  or     DeleteFilesOlderThan /t=<time specification> /p=<folder to delete files from> [/f=<files to delete>[;<files to delete>]] [/s=0|1]"
   Con.StdOut "i.e.     DeleteFilesOlderThan /time=2d /path=D:\MyTarget"
   Con.StdOut "         DeleteFilesOlderThan /t=3w /p=C:\MyTarget\Data /f=*.txt /s=1"
   Con.StdOut ""
   Con.StdOut "Parameters"
   Con.StdOut "----------"
   Con.StdOut "/t or /time        = time specification"
   Con.StdOut "/p or /path        = (start) folder"
   Con.StdOut "/f or /filepattern = file pattern"
   Con.StdOut "       If omitted, all files are scanned (equals /f=*.*)."
   Con.StdOut "/s or /subfolders  = recurse subfolders yes(1) or no (0)"
   Con.StdOut "       If omitted, only the folder passed via /p is scanned for matching files (equals /s=0)."
   Con.StdOut ""
   Con.StdOut "You may specify more than one file pattern for the parameter /f by using ; (semicolon) as a separator, i.e."
   Con.StdOut "       /f=*.doc;*.rtf -> deletes all *.doc and all *.rtf files from the specified folder."
   Con.StdOut "       /f=Backup*.bak;Log*.trn -> deletes all Backup*.bak and all Log*.trn files from the specified folder."
   Con.StdOut ""
   Con.StdOut "Allowed time specification units for parameter /t are:
   Con.StdOut "        d = day   i.e. 1d"
   Con.StdOut "        w = week  i.e. 2w"
   Con.StdOut "        m = month i.e. 3m"
   Con.StdOut "        y = year  i.e. 4y"
   Con.StdOut ""

End Sub
'---------------------------------------------------------------------------
