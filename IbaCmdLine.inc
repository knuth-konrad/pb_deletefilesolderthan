'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
$DELIM_PARAM = "/"
$DELIM_VALUE = "="

Class cBACmdLine


   Instance msParamDelimiter As String ' Parameterbegrenzer, üblicherweise ein "/"
   Instance msValueDelimiter As String ' Begrenzer zwischen Parametername und
                                       ' Parameterwert, üblicherweise ein "="
   Instance mdwParamCount As Dword     ' Anzahl der Parameter
   Instance masParams() As String      ' Alle Parameternamen
   Instance mavntValues() As Variant   ' Alle Parameterwerte. Variant weil ein Wert
                                       ' alles sein kann

   Instance mdwParamID As Dword        ' Collection key of parameter
   Instance mdwValueID As Dword        ' Collection key of value

   Instance mcolValues As IPowerCollection   ' Alle Parameterwerte. Variant weil ein Wert
                                             ' alles sein kann
'------------------------------------------------------------------------------

   Class Method Create()
   ' Do initialization

      ' Set defaults
      msParamDelimiter = $DELIM_PARAM
      msValueDelimiter = $DELIM_VALUE

      ' Initialize collections
      Let mcolValues = Class "PowerCollection"

   End Method
'------------------------------------------------------------------------------

   Class Method Destroy()
   ' Do cleanup

      ' Release resources
      Let mcolValues = Nothing

   End Method
'------------------------------------------------------------------------------

   Interface IBACmdLine

      Inherit Dual

      ' ParamDelimiter
      Property Get ParamDelimiter() As String
         Property = msParamDelimiter
      End Property
      Property Set ParamDelimiter(ByVal sValue As String)
         msParamDelimiter = sValue
      End Property
'------------------------------------------------------------------------------
      ' ValueDelimiter
      Property Get ValueDelimiter() As String
         Property = msValueDelimiter
      End Property
      Property Set ValueDelimiter(ByVal sValue As String)
         msValueDelimiter = sValue
      End Property
'------------------------------------------------------------------------------

      Method ErrString(ByVal lErr As Long, Optional ByVal vntPrefix As Variant) As String
      '------------------------------------------------------------------------------
      'Purpose  : Returns an formatted error string from an (PB) error number
      '
      'Prereq.  : -
      'Parameter: -
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local sPrefix As String

         If Not IsMissing(vntPrefix) Then
            sPrefix = Variant$(vntPrefix)
         End If

         Method = sPrefix & Format$(lErr) & " - " & Error$(lErr)

      End Method
      '------------------------------------------------------------------------------

      Method GetParamByIndex(ByVal lIndex As Long) As String

         Local vntValue As Variant

         Trace On
         Trace Print FuncName$

         ' *** Safe guard
         If lIndex < 1 Or lIndex > Me.ValuesCount Then
            Method = ""
            Exit Method
         End If

         Local wsKey As WString

         Try
            mcolValues.Entry lIndex, wsKey, vntValue

            If ObjResult = %S_False Then
               Method = ""
               Trace Print " -- GetParamByIndex: ObjResult = %S_FALSE"
            Else
               Method = wsKey
            End If
         Catch
            Trace Print " -- GetParamByIndex (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method HasParam(ByVal wsKey As WString, Optional ByVal vntKeyAlias As Variant) As Long
      '------------------------------------------------------------------------------
      'Purpose  : Check wether a specific command line parameter is present
      '
      'Prereq.  : -
      'Parameter: wsKey       - parameter to search for
      '           vntKeyAlias - parameter alias
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad
      '   Source: -
      '  Changed: 10.11.2016
      '           - Check for an additional (alias) parameter having the same
      '           meaning at the same time.
      '           I.e. /f and /file
      '------------------------------------------------------------------------------
         Local vntValue As Variant

         Trace On
         Trace Print FuncName$

         ' *** Safe guard
         If (Len(wsKey) < 1) And (IsMissing(vntKeyAlias)) Then
            Method = %TRUE
            Exit Method
         End If

         Try
            vntValue = mcolValues.Item(wsKey)

            If ObjResult = %S_False Then
               Method = %FALSE
               Trace Print " -- HasParam wsKey: ObjResult = %S_FALSE"
               Exit Method
            Else
               Method = %TRUE
               Exit Method
            End If
         Catch
            Trace Print " -- HasParam (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

         If Not IsMissing(vntKeyAlias) Then
            Local wsKeyAlias As WString
            wsKeyAlias = Variant$$(vntKeyAlias)

            Try
               vntValue = mcolValues.Item(wsKeyAlias)

               If ObjResult = %S_False Then
                  Method = %FALSE
                  Trace Print " -- HasParam vntKeyAlias: ObjResult = %S_FALSE"
                  Exit Method
               Else
                  Method = %TRUE
                  Exit Method
               End If
            Catch
               Trace Print " -- HasParam (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
               ErrClear
            End Try

         End If

      End Method
      '------------------------------------------------------------------------------

      Method ValuesAdd(ByVal wsKey As WString, ByVal vntValue As Variant) As String

         Local hResult As Long

         Trace On
         Trace Print FuncName$

         Trace Print "  - wsKey: " & wsKey
         Trace Print "  - vntValue: " & Variant$(vntValue)

         Try
            mcolValues.Add(wsKey, vntValue)
            hResult = ObjResult
            Method = wsKey
         Catch
            Trace Print " -- ValuesAdd (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesClear()

         Trace On
         Trace Print FuncName$

         Try
            mcolValues.Clear()
         Catch
            Trace Print " -- ValuesClear (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesCount() As Dword

         Trace On
         Trace Print FuncName$

         Try
            Method = mcolValues.Count
         Catch
            Trace Print " -- ValuesCount Error: " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
            Method = 0
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesGet(ByVal lIndex As Long) As Variant

         Local vntValue As Variant, wsKey As WString

         Trace On
         Trace Print FuncName$

         Try
            If lIndex >= 1 And lIndex <= mcolValues.Count Then
               mcolValues.Entry lIndex, wsKey, vntValue
               If ObjResult = %S_False Then
                  Trace Print " -- ValuesGet: ObjResult = %S_FALSE"
               Else
                  Method = vntValue
               End If
            End If
         Catch
            Trace Print " -- ValuesGet(lIndex) Error: (" & Format$(lIndex) & ") " & _
               Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method GetValueByIndex(ByVal lIndex As Long) As Variant
      '------------------------------------------------------------------------------
      'Purpose  : Liefert den Wert eines Parameters
      '
      'Prereq.  : -
      'Parameter: vntIndex -  Nummer des Parameters (1 bis Anzahl Parameter(=ParamCount))
      'Returns  : Wert des Parameters
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

         Trace On
         Trace Print FuncName$

         ' Index im gültigen Bereich?
         If lIndex >= 1 And lIndex <= mcolValues.Count Then
            Method = me.ValuesGet(lIndex)
         End If

      End Method
      '------------------------------------------------------------------------------

      Method GetValueByName(ByVal sParam As String, Optional ByVal vntParamAlias As Variant,  _
         Optional vntCaseSensitive As Variant) As Variant
      '------------------------------------------------------------------------------
      'Purpose  : Liefert den Wert eines Parameters
      '
      'Prereq.  : -
      'Parameter: sParam            -  Name des Parameters dessen Wert gesucht wird
      '           bolCaseSensitive  -  Groß-/Kleinschreibung beachten
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: 10.11.2016
      '           - Check for an additional (alias) parameter having the same
      '           meaning at the same time.
      '           I.e. /f and /file
      '------------------------------------------------------------------------------
         Local bolCaseSensitive As Long
         Local vntValue As Variant
         Local wsParam, wsParamAlias As WString

         Trace On
         Trace Print FuncName$

         wsParam = sParam

         If IsMissing(vntCaseSensitive) Then
            bolCaseSensitive = %FALSE
         Else
            If IsTrue(Variant#(vntCaseSensitive)) Then
               bolCaseSensitive = %TRUE
            Else
               bolCaseSensitive = %FALSE
            End If
         End If

         If IsTrue(bolCaseSensitive) Then

            Local i As Dword
            Local wsKey As WString

            For i = 1 To mcolValues.Count
               mcolValues.Entry i, wsKey, vntValue
               If LCase$(wsParam) = LCase$(wsKey) Then
                  Method = vntValue
                  Exit Method
               End If
            Next i

            If Not IsMissing(vntParamAlias) Then
               wsParamAlias = Variant$$(vntParamAlias)

               For i = 1 To mcolValues.Count
                  mcolValues.Entry i, wsKey, vntValue
                  If LCase$(wsParamAlias) = LCase$(wsKey) Then
                     Method = vntValue
                     Exit Method
                  End If
               Next i
            End If

         Else

            Try
               Method = mcolValues.Item(wsParam)
               Exit Method
            Catch
               Trace Print " -- GetValueByName() Error: " & Format$(Err) & ", " & Me.ErrString(Err)
               ErrClear
            End Try

            If Not IsMissing(vntParamAlias) Then
               wsParamAlias = Variant$$(vntParamAlias)

               Try
                  Method = mcolValues.Item(wsParamAlias)
                  Exit Method
               Catch
                  Trace Print " -- GetValueByName() Error: " & Format$(Err) & ", " & Me.ErrString(Err)
                  ErrClear
               End Try

            End If

         End If

      End Method
      '------------------------------------------------------------------------------

      Method Init(ByVal sCmd As String) As Long
      '------------------------------------------------------------------------------
      'Purpose  : Diese Prozedur muß els erste Prozedur aufgerufen werden. Sie übernimmt
      '           die Kommandozeile und wertet sie aus.
      '
      'Prereq.  : -
      'Parameter: sCmd  -  Per COMMAND$ empfangene Kommandozeile
      'Returns  : True  -  Kommandozeile konnte ausgewertet werden
      '           False -  Ein Fehler ist aufgetreten
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local i, lParamCount As Long
         Local vntValue As Variant
         Dim awsParams() As WString

         Trace On
         Trace Print FuncName$

         Trace Print "  - sCmd: " & sCmd

         On Error GoTo InitError

         ' ** Safe guards
         ' Enthielt CLI Werte?
         sCmd = LTrim$(sCmd)
         lParamCount = ParseCount(sCmd, msParamDelimiter)
         If lParamCount < 2 Then
            Method = %TRUE
            Exit Method
         End If

         Trace Print "  - ParseCount(sCmd): " & Format$(lParamCount)

         ' ParseCount returns 3 for strings like "/123 /abc" where delimiter is '/', as there's an
         ' 'empty' entry in front of the first '/'
         Dim awsParams(lParamCount - 1) As WString
         Parse sCmd, awsParams(), msParamDelimiter
         Me.ValuesClear

         Trace Print "  - LBound(awsParams): " & Format$(LBound(awsParams))
         Trace Print "  - UBound(awsParams): " & Format$(UBound(awsParams))

         For i = LBound(awsParams) To UBound(awsParams)
            Local wsKey As WString

            Trace Print "  - awsParams(i): (" & Format$(i) & "), " & awsParams(i)

            ' Nur wenn auch ein Parameter da ist...
            If Len(Trim$(awsParams(i))) > 0 Then

               wsKey = Trim$(Remove$(Extract$(awsParams(i), msValueDelimiter), msParamDelimiter))

               Trace Print "  - wsKey: " & wsKey

               ' Parameter in der Art /User=Knuth.
               ' "User" ist der Parameter, "Knuth" ist der Wert
               If InStr(awsParams(i), msValueDelimiter) > 0 Then
                  vntValue = Trim$(Remain$(awsParams(i), msValueDelimiter))

                  Trace Print "  - vntValue: " & Variant$(vntValue)

                  Me.ValuesAdd(wsKey, vntValue)

               Else
               'Parameter in der Art /Quit.
               '"Quit" ist der Parameter, als Wert wird "True" angenommen
                  Me.ValuesAdd(wsKey, %TRUE)

                  Trace Print "  - vntValue: %TRUE"

               End If

            End If

         Next i

         Method = %TRUE

      InitExit:
         On Error GoTo 0
         Exit Method

      InitError:
         Method = %FALSE
         Trace Print " -- Init(sCmd) Error: " & sCmd & Format$(Err) & ", " & Me.ErrString(Err)
         ' Collection "aufräumen"
         Me.ValuesClear
         ErrClear
         Resume InitExit

      End Method
      '==============================================================================

   End Interface
'------------------------------------------------------------------------------

End Class
