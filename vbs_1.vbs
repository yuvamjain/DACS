'----------------------------------------------------------------------------------------------------------------------------
'Script Name : QueryEventLogs.vbs
'Author      : Matthew Beattie
'Created     : 13/10/09
'Description : This script queries the event log for...whatever you want it to! Just set the event log name and event ID's!
'----------------------------------------------------------------------------------------------------------------------------
'Initialization  Section
'----------------------------------------------------------------------------------------------------------------------------
Option Explicit
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Dim objDictionary, objFSO, wshShell, wshNetwork
Dim scriptBaseName, scriptPath, scriptLogPath
Dim ipAddress, macAddress, item, messageType, message
On Error Resume Next
   Set objDictionary = NewDictionary
   Set objFSO        = CreateObject("Scripting.FileSystemObject")
   Set wshShell      = CreateObject("Wscript.Shell")
   Set wshNetwork    = CreateObject("Wscript.Network")
   scriptBaseName    = objFSO.GetBaseName(Wscript.ScriptFullName)
   scriptPath        = objFSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
   scriptLogPath     = scriptPath & "\" & IsoDateString(Now)
   If Err.Number <> 0 Then
      Wscript.Quit
   End If
On Error Goto 0
'----------------------------------------------------------------------------------------------------------------------------
'Main Processing Section
'----------------------------------------------------------------------------------------------------------------------------
On Error Resume Next
   PromptScriptStart
   ProcessScript
   If Err.Number <> 0 Then
      MsgBox BuildError("Processing Script"), vbCritical, scriptBaseName
      Wscript.Quit
   End If
   PromptScriptEnd
On Error Goto 0
'----------------------------------------------------------------------------------------------------------------------------
'Functions Processing Section
'----------------------------------------------------------------------------------------------------------------------------
'Name       : ProcessScript -> Primary Function that controls all other script processing.
'Parameters : None          ->
'Return     : None          ->
'----------------------------------------------------------------------------------------------------------------------------
Function ProcessScript
   Dim hostName, logName, startDateTime, endDateTime
   Dim events, eventNumbers, i
   hostName      = wshNetwork.ComputerName
   logName       = "Security"
   eventNumbers  = Array("672")
   startDateTime = DateAdd("n", -120, Now)
   '-------------------------------------------------------------------------------------------------------------------------
   'Query the event log for the eventID's within the specified event log name and date range.
   '-------------------------------------------------------------------------------------------------------------------------
   If Not QueryEventLog(events, hostName, logName, eventNumbers, startDateTime) Then
      Exit Function
   End If
   '-------------------------------------------------------------------------------------------------------------------------
   'Log the scripts results to the scripts
   '-------------------------------------------------------------------------------------------------------------------------
   For i = 0 To UBound(events)
      LogMessage events(i)
   Next
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : QueryEventLog -> Primary Function that controls all other script processing.
'Parameters : results       -> Input/Output : Variable assigned to an array of results from querying the event log.
'           : hostName      -> String containing the hostName of the system to query the event log on.
'           : logName       -> String containing the name of the Event Log to query on the system.
'           : eventNumbers  -> Array containing the EventID's (eventCode) to search for within the event log.
'           : startDateTime -> Date\Time containing the date to finish searching at.
'           : minutes       -> Integer containing the number of minutes to subtract from the startDate to begin the search.
'Return     : QueryEventLog -> Returns True if the event log was successfully queried otherwise returns False.
'----------------------------------------------------------------------------------------------------------------------------
Function QueryEventLog(results, hostName, logName, eventNumbers, startDateTime)
   Dim wmiDateTime, wmi, query, eventItems, eventItem
   Dim timeWritten, eventDate, eventTime, description
   Dim eventsDict, eventInfo, errorCount, i
   QueryEventLog = False
   errorCount    = 0
   If Not IsArray(eventNumbers) Then
      eventNumbers = Array(eventNumbers)
   End If
   '-------------------------------------------------------------------------------------------------------------------------
   'Construct part of the WMI Query to account for searching multiple eventID's
   '-------------------------------------------------------------------------------------------------------------------------
   query = "Select * from Win32_NTLogEvent Where Logfile = " & SQ(logName) & " And (EventCode = "
   For i = 0 To UBound(eventNumbers)
      query = query & SQ(eventNumbers(i)) & " Or EventCode = "
   Next
   On Error Resume Next
      Set eventsDict = NewDictionary
      If Err.Number <> 0 Then
         LogError "Creating Dictionary Object"
         Exit Function
      End If
      Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & hostName & "\root\cimv2")
      If Err.Number <> 0 Then
         LogError "Creating WMI Object to connect to " & DQ(hostName)
         Exit Function
      End If
      '----------------------------------------------------------------------------------------------------------------------
      'Create the "SWbemDateTime" Object for converting WMI Date formats. Supported in Windows Server 2003 & Windows XP.
      '----------------------------------------------------------------------------------------------------------------------
      Set wmiDateTime = CreateObject("WbemScripting.SWbemDateTime")
      If Err.Number <> 0 Then
         LogError "Creating " & DQ("WbemScripting.SWbemDateTime") & " object"
         Exit Function
      End If
      '----------------------------------------------------------------------------------------------------------------------
      'Build the WQL query and execute it.
      '----------------------------------------------------------------------------------------------------------------------
      wmiDateTime.SetVarDate startDateTime, True
      query          = Left(query, InStrRev(query, "'")) & ") And (TimeWritten >= " & SQ(wmiDateTime.Value) & ")"
      Set eventItems = wmi.ExecQuery(query)
      If Err.Number <> 0 Then
         LogError "Executing WMI Query " & DQ(query)
         Exit Function
      End If
      '----------------------------------------------------------------------------------------------------------------------
      'Convert the property values of Each event found to a comma seperated string and add it to the dictionary.
      '----------------------------------------------------------------------------------------------------------------------
      For Each eventItem In eventItems
         Do
            timeWritten = ""
            eventDate   = ""
            eventTime   = ""
            eventInfo   = ""
            timeWritten = ConvertWMIDateTime(eventItem.TimeWritten)
            eventDate   = FormatDateTime(timeWritten, vbShortDate)
            eventTime   = FormatDateTime(timeWritten, vbLongTime)
            eventInfo   = eventDate                          & ","
            eventInfo   = eventInfo & eventTime              & ","
            eventInfo   = eventInfo & eventItem.SourceName   & ","
            eventInfo   = eventInfo & eventItem.Type         & ","
            eventInfo   = eventInfo & eventItem.Category     & ","
            eventInfo   = eventInfo & eventItem.EventCode    & ","
            eventInfo   = eventInfo & eventItem.User         & ","
            eventInfo   = eventInfo & eventItem.ComputerName & ","
            description = eventItem.Message
            '------------------------------------------------------------------------------------------------------------------------
            'Ensure the event description is not blank.
            '------------------------------------------------------------------------------------------------------------------------
            If IsNull(description) Then
               description = "The event description cannot be found."
            End If
            description = Replace(description, vbCrLf, " ")
            eventInfo   = eventInfo & description
            '------------------------------------------------------------------------------------------------------------------------
            'Check if any errors occurred enumerating the event Information
            '------------------------------------------------------------------------------------------------------------------------
            If Err.Number <> 0 Then
               LogError "Enumerating Event Properties from the " & DQ(logName) & " event log on " & DQ(hostName)
               errorCount = errorCount + 1
               Err.Clear
               Exit Do
            End If
            '------------------------------------------------------------------------------------------------------------------------
            'Remove all Tabs and spaces.
            '------------------------------------------------------------------------------------------------------------------------
            eventInfo = Trim(Replace(eventInfo, vbTab, " "))
            Do While InStr(1, eventInfo, "  ", vbTextCompare) <> 0
               eventInfo = Replace(eventInfo, "  ", " ")
            Loop
            '------------------------------------------------------------------------------------------------------------------------
            'Add the Event Information to the Dictionary object if it doesn't exist.
            '------------------------------------------------------------------------------------------------------------------------
            If Not eventsDict.Exists(eventInfo) Then
               eventsDict(eventsDict.Count) = eventInfo
            End If
         Loop Until True
      Next
   On Error Goto 0
   If errorCount <> 0 Then
      Exit Function
   End If
   results       = eventsDict.Items
   QueryEventLog = True
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : ConvertWMIDateTime -> Converts a WMI Date Time String into a String that can be formatted as a valid Date Time.
'Parameters : wmiDateTimeString  -> String containing a WMI Date Time String.
'Return     : ConvertWMIDateTime -> Returns a valid Date Time String otherwise returns a Blank String.
'----------------------------------------------------------------------------------------------------------------------------
Function ConvertWMIDateTime(wmiDateTimeString)
   Dim integerValues, i
   '-------------------------------------------------------------------------------------------------------------------------
   'Ensure the wmiDateTimeString contains a "+" or "-" character. If it doesn't it is not a valid WMI date time so exit.
   '-------------------------------------------------------------------------------------------------------------------------
   If InStr(1, wmiDateTimeString, "+", vbTextCompare) = 0 And _
      InStr(1, wmiDateTimeString, "-", vbTextCompare) = 0 Then
      ConvertWMIDateTime = ""
      Exit Function
   End If
   '-------------------------------------------------------------------------------------------------------------------------
   'Replace any "." or "+" or "-" characters in the wmiDateTimeString and check each character is a valid integer.
   '-------------------------------------------------------------------------------------------------------------------------   
   integerValues = Replace(Replace(Replace(wmiDateTimeString, ".", ""), "+", ""), "-", "")
   For i = 1 To Len(integerValues)
      If Not IsNumeric(Mid(integerValues, i, 1)) Then
         ConvertWMIDateTime = ""
         Exit Function
      End If
   Next
   '-------------------------------------------------------------------------------------------------------------------------
   'Convert the WMI Date Time string to a String that can be formatted as a valid Date Time value.
   '-------------------------------------------------------------------------------------------------------------------------
   ConvertWMIDateTime = CDate(Mid(wmiDateTimeString, 5, 2)  & "/" & _
                              Mid(wmiDateTimeString, 7, 2)  & "/" & Left(wmiDateTimeString, 4) & " " & _
                              Mid(wmiDateTimeString, 9, 2)  & ":" & _
                              Mid(wmiDateTimeString, 11, 2) & ":" & _
                              Mid(wmiDateTimeString, 13, 2))
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : NewDictionary -> Creates a new dictionary object.
'Parameters : None          ->
'Return     : NewDictionary -> Returns a dictionary object.
'----------------------------------------------------------------------------------------------------------------------------
Function NewDictionary
   Dim dict
   Set dict          = CreateObject("scripting.Dictionary")
   dict.CompareMode  = vbTextCompare
   Set NewDictionary = dict
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : SQ          -> Places single quotes around a string
'Parameters : stringValue -> String containing the value to place single quotes around
'Return     : SQ          -> Returns a single quoted string
'----------------------------------------------------------------------------------------------------------------------------
Function SQ(ByVal stringValue)
   If VarType(stringValue) = vbString Then
      SQ = "'" & stringValue & "'"
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : DQ          -> Place double quotes around a string and replace double quotes
'           :             -> within the string with pairs of double quotes.
'Parameters : stringValue -> String value to be double quoted
'Return     : DQ          -> Double quoted string.
'----------------------------------------------------------------------------------------------------------------------------
Function DQ (ByVal stringValue)
   If stringValue <> "" Then
      DQ = """" & Replace (stringValue, """", """""") & """"
   Else
      DQ = """"""
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : IsoDateTimeString -> Generate an ISO date and time string from a date/time value.
'Parameters : dateValue         -> Input date/time value.
'Return     : IsoDateTimeString -> Date and time parts of the input value in "yyyy-mm-dd hh:mm:ss" format.
'----------------------------------------------------------------------------------------------------------------------------
Function IsoDateTimeString(dateValue)
   IsoDateTimeString = IsoDateString (dateValue) & " " & IsoTimeString (dateValue)
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : IsoDateString -> Generate an ISO date string from a date/time value.
'Parameters : dateValue     -> Input date/time value.
'Return     : IsoDateString -> Date part of the input value in "yyyy-mm-dd" format.
'----------------------------------------------------------------------------------------------------------------------------
Function IsoDateString(dateValue)
   If IsDate(dateValue) Then
      IsoDateString = Right ("000" &  Year (dateValue), 4) & "-" & _
                      Right (  "0" & Month (dateValue), 2) & "-" & _
                      Right (  "0" &   Day (dateValue), 2)
   Else
      IsoDateString = "0000-00-00"
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : IsoTimeString -> Generate an ISO time string from a date/time value.
'Parameters : dateValue     -> Input date/time value.
'Return     : IsoTimeString -> Time part of the input value in "hh:mm:ss" format.
'----------------------------------------------------------------------------------------------------------------------------
Function IsoTimeString(dateValue)
   If IsDate(dateValue) Then
      IsoTimeString = Right ("0" &   Hour (dateValue), 2) & ":" & _
                      Right ("0" & Minute (dateValue), 2) & ":" & _
                      Right ("0" & Second (dateValue), 2)
   Else
      IsoTimeString = "00:00:00"
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : LogMessage -> Writes a message to a log file.
'Parameters : logPath    -> String containing the full folder path and file name of the Log file without with file extension.
'           : message    -> String containing the message to include in the log message.
'Return     : None       -> 
'----------------------------------------------------------------------------------------------------------------------------
Function LogMessage(message)
   If Not LogToCentralFile(scriptLogPath & ".log", IsoDateTimeString(Now) & "," & message) Then
      Exit Function
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : LogError -> Writes an error message to a log file.
'Parameters : logPath  -> String containing the full folder path and file name of the Log file without with file extension.
'           : message  -> String containing a description of the event that caused the error to occur.
'Return     : None       -> 
'----------------------------------------------------------------------------------------------------------------------------
Function LogError(message)
   If Not LogToCentralFile(scriptLogPath & ".err", IsoDateTimeString(Now) & "," & BuildError(message)) Then
      Exit Function
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name      : BuildError -> Builds a string of information relating to the error object.
'Parameters: message    -> String containnig the message that relates to the process that caused the error.
'Return    : BuildError -> Returns a string relating to error object.   
'----------------------------------------------------------------------------------------------------------------------------
Function BuildError(message)
   BuildError = "Error " & Err.Number & " (Hex " & Hex(Err.Number) & ") " & message & ". " & Err.Description
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : LogToCentralFile -> Attempts to Appends information to a central file.
'Parameters : logSpec          -> Folder path, file name and extension of the central log file to append to.
'           : message          -> String to include in the central log file
'Return     : LogToCentralFile -> Returns True if Successfull otherwise False.
'----------------------------------------------------------------------------------------------------------------------------
Function LogToCentralFile(logSpec, message)
   Dim attempts, objLogFile
   LogToCentralFile = False
   '-------------------------------------------------------------------------------------------------------------------------
   'Attempt to append to the central log file up to 10 times, as it may be locked by some other system.
   '-------------------------------------------------------------------------------------------------------------------------
   attempts = 0
   Do
      On Error Resume Next
         Set objLogFile = objFSO.OpenTextFile(logSpec, ForAppending, True)
         If Err.Number = 0 Then
            objLogFile.WriteLine message
            objLogFile.Close
            LogToCentralFile = True
            Exit Function
         End If
      On Error Goto 0
      Randomize
      Wscript.sleep 1000 + Rnd * 100
      attempts = attempts + 1
   Loop Until attempts >= 10
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : PromptScriptStart -> Prompt when script starts.
'Parameters : None
'Return     : None
'----------------------------------------------------------------------------------------------------------------------------
Function PromptScriptStart
   MsgBox "Now processing the " & DQ(Wscript.ScriptName) & " script.", vbInformation, scriptBaseName
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : PromptScriptEnd -> Prompt when script has completed.
'Parameters : None
'Return     : None
'----------------------------------------------------------------------------------------------------------------------------
Function PromptScriptEnd
   MsgBox "The " & DQ(Wscript.ScriptName) & " script has completed successfully.", vbInformation, scriptBaseName
End Function
'------------------------------------------------------------