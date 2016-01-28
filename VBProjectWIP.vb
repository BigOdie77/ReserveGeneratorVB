'
'Program Name: RsvGeneratorVB
'Purpose: Reads in all the Panels and corresponding information from a text file.
'         Reads in the Schedule information and loads it to a variable from a text file.
'         Writes the schedule to the database.
'         Creates an error log containing panels that were not responding and rooms that do not exist on the system
'         Creates an updated version of the panel look up table containing just the panels that were not responding the first time through.
'         Creates a note file containing notes submitted from the user when creating a schedule.  
'         
'Date: Janurary 13th, 2016
'Author: Eric Odette
'

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class Form1
    Public RevGeneratorVB()

    'Boolean Flags
    Private genFlag As Boolean = False
    Private ignoreFlag As Boolean = False
    Private firstRun As Boolean = False
    Private invalidFlag As Boolean = False

    'Private Const TWELVE_TF As DateFormat = New DateFormat("h:mm tt")
    Private TWELVE_TF As String = "hh:mma"
    Private TWENTY_FOUR_TF As String = "{HH:mm:ss}"

    Private finalEnd, finalStart, room, desc, instanceStr, devIDStr, panel As String

    'Database Variables
    Dim Cnxn As New ADODB.Connection
    Dim cmdChange As ADODB.Command
    Dim Result As ADODB.Recordset
    Dim conn As String = "Provider=MSDASQL.1; Password=LOGIN; User ID=DELTA;Data Source=Delta Network" 'Dsn=Delta Network;uid=DELTA
    Dim cmdStr As String

    Private day As Integer
    Private progress As Double = 0.0
    Private fileSize As Long
    Private panelLine As String = ""

    'No hashtables in vb use Dictionary instead or something like it
    Dim roomMap As New Dictionary(Of String, String)
    Dim missedScheduleMap As New Dictionary(Of String, String)
    Dim unfoundRoomMap As New Dictionary(Of String, String)

    'ArrayLists
    Private tempPanelInfoArray As New ArrayList() 'Type PanelData
    Private lookupArray As New ArrayList()        'Type String

    'PanelData Class
    '   Holds data related to the panels like the device Id, instance, start time and end time
    Public Class PanelData
        Public dev As String = ""
        Public inst As String = ""
        Public tempStart As Double = 0.0
        Public tempEnd As Double = 0.0
    End Class

    '
    'Method Name: functionWorker()
    'Purpose: To run all the methods for generating the schedules
    '
    Public Sub functionWorker()
        Try
            defaultSettings()
            btnGenSch.Enabled = False
            dbConnect()
            readFile(TextBoxBrowseSch.Text)
            dbClose()
            btnGenSch.Enabled = True
            progBar.Value = progBar.Maximum
            TextBoxErrorLog.Text = "Error Log: "
            writeErrorLog()
            writeNewLookup()

            If missedScheduleMap.Count = 0 And unfoundRoomMap.Count = 0 Then
                MsgBox("Schedules Created!")
            ElseIf Not missedScheduleMap.Count = 0 And unfoundRoomMap.Count = 0 Then
                MsgBox("Schedules Created with some errors." & vbNewLine & "Please veiw the Error Log below for further detail")
            Else
                MsgBox("Schedules Created with some errors." & vbNewLine & "Please veiw the Error Log below for further detail")
                btnGenMissedSch.Enabled = True
            End If
            missedScheduleMap.Clear()

        Catch ex As Exception
            MsgBox("ERROR in generating schedule: " & ex.Message)
            'TextBoxErrorLog.Text = " Error Log: "
            progBar.Value = 0
            dbClose()
        End Try
    End Sub

    '
    'Method Name: dbConnect()
    'Purpose: Connects to the database using the Delta Network obdc data source
    '
    Private Sub dbConnect()
        Try
            TextBoxErrorLog.AppendText(vbNewLine & "Status: Connecting to Database... " & vbNewLine)
            Cnxn = New ADODB.Connection
            Cnxn.ConnectionString = "Dsn=Test Panels;uid=DELTA"
            Cnxn.Open()
            TextBoxErrorLog.AppendText(vbNewLine & "Status: Database connected..." & vbNewLine)
        Catch ex As SqlException
            System.Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Sub ' end dbConnect

    '
    ' Method Name: dbClose() 
    ' Purpose: Used to close the database connection to prevent a hanging connection
    '
    Private Sub dbClose()
        TextBoxErrorLog.AppendText(vbNewLine & "Status: Closing database connection..." & vbNewLine)
        Try
            If Cnxn IsNot Nothing Then
                Cnxn.Close()
            End If

            TextBoxErrorLog.AppendText(vbNewLine & "Status: Database Closed." & vbNewLine)
        Catch ex As SqlException
            System.Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: readFile()
    'Purpose: reads the lines of a given file and performs regular expression search for the 
    '         necessary information
    '
    ' * Check Calendar to make sure its working right ... *
    '
    Public Sub readFile(ByRef filePath As String)

        Dim tempDay As Integer = 0
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(filePath)

        Dim strLine As String = fileReader.ReadLine()
        Dim dayStr = strLine.Substring(0, 50).Trim()
        dayStr = dayStr.Substring(0, 9)

        Dim pattern As String = "((Monday)|(Tuesday)|(Wednesday)|(Thursday)|(Friday)|(Saturday)|(Sunday))"

        'For loop to grab the day of the week
        If dayStr.Contains(",") Then
            For ind As Integer = 0 To dayStr.Length()
                If dayStr.ElementAt(ind) = "," Then
                    dayStr = dayStr.Substring(0, ind)
                    Exit For
                End If
            Next
        End If

        While Not fileReader.EndOfStream 'Line is blank, skip over it
            If strLine.Equals("") Or strLine.Equals("vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ") Then
                While strLine.Equals("") And Not fileReader.EndOfStream
                    strLine = fileReader.ReadLine()
                End While
            End If

            Dim matcher As MatchCollection = Regex.Matches(dayStr, pattern)
            For idx As Integer = 0 To matcher.Count()
                Dim matchFound As Boolean = matcher.Item(idx).Success

                If matchFound Then
                    If matcher.Item(0).ToString.Equals("Monday") Then
                        day = 1
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Tuesday") Then
                        day = 2
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Wednesday") Then
                        day = 3
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Thursday") Then
                        day = 4
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Friday") Then
                        day = 5
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Saturday") Then
                        day = 6
                        Exit For
                    ElseIf matcher.Item(0).ToString.Equals("Sunday") Then
                        day = 7
                        Exit For
                    End If
                End If
            Next

            If day <> tempDay Then
                TextBoxErrorLog.AppendText(vbNewLine & "Status: Clearing Schedules... " & vbNewLine)
                If Not btnGenMissedSch.Enabled() Then
                    clearSchedule(TextBoxBrowseLookup.Text)
                ElseIf btnGenMissedSch.Enabled Then
                    clearSchedule(TextBoxGenMissedSch.Text)
                End If
                tempDay = day
            End If

            strLine = strLine.Trim()

            Dim values As String
            values = strLine.Substring(56, 85).Trim()
            Dim startTime, endTime

            startTime = values.Substring(0, 7).Trim()

            Dim startTimeTemp As DateTime = startTime
            startTimeTemp = startTimeTemp.AddHours(-1)
            startTime = startTimeTemp

            endTime = values.Substring(9, 16).Trim()

            finalStart = convertTo24HoursFormat(startTime)
            finalEnd = convertTo24HoursFormat(endTime)

            values = values.Substring(17).Trim()
            room = values.Substring(0, 25).Trim()

            values = values.Substring(25).Trim()
            desc = values.Trim()


            If Not checkDesc(desc) Then
                'do not use this room. no need to allocate a schedule
            Else
                TextBoxErrorLog.AppendText(vbNewLine & "Status: Creating new schedules..." & vbNewLine)
                If Not btnGenMissedSch.Enabled() Then
                    readLookupTable(TextBoxBrowseLookup.Text)
                ElseIf btnGenMissedSch.Enabled() Then
                    readLookupTable(TextBoxGenMissedSch.Text)
                End If

                progress += 1 'More strLine.getBytes().length
                progBar.Value = progress
            End If
            strLine = fileReader.ReadLine()
        End While
    End Sub

    '
    'Method Name: readLookupTable()
    'Purpose: performs a regular expression search on the look up table to get panel information
    '
    Public Sub readLookupTable(s As String)
        Try
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(s)
            Dim strLine As String = fileReader.ReadLine()

            ' found flag is used so that if the room for that schedule is not found in
            ' the Lookup table it will be put to the map to be displayed later
            Dim foundFlag As Boolean = False

            'Dictionary/ Hashmap to hold the new missing rooms or schedules
            Dim ignoreRoomMap As New Dictionary(Of String, String)

            While Not fileReader.EndOfStream
                If ignoreFlag = False Then
                    Dim lines As String = strLine.Substring(47)
                    Dim ignoreRooms() As String = lines.Split(", ")

                    For index As Integer = 0 To ignoreRooms.Length() - 1
                        ignoreRoomMap.Add(ignoreRooms.ElementAt(index), ignoreRooms.ElementAt(index))
                    Next
                    ignoreFlag = True

                ElseIf ignoreFlag = True Then
                    panelLine = strLine
                    roomMap.Clear()

                    'While loop to go through text until its got panel info 
                    While Not strLine.Contains("###")
                        strLine = fileReader.ReadLine()
                    End While

                    Dim currentPanelInfoArray As String()
                    currentPanelInfoArray = strLine.Split("#")

                    If strLine.Contains("##") Then
                        panel = currentPanelInfoArray.ElementAt(0)
                        Dim roomList As String = currentPanelInfoArray.ElementAt(3)
                        devIDStr = currentPanelInfoArray.ElementAt(6)
                        instanceStr = currentPanelInfoArray.ElementAt(9)

                        Dim rooms() As String = roomList.Split(", ")
                        For i As Integer = 0 To rooms.Length() - 1
                            roomMap.Add(rooms(i), rooms(i))
                        Next

                    End If
                    strLine = fileReader.ReadLine()

                    Dim parsedStart As Double = Double.Parse(finalStart.Substring(0, 5).Replace(":", "."))
                    Dim parsedEnd As Double = Double.Parse(finalEnd.Substring(0, 5).Replace(":", "."))

                    If roomMap.ContainsValue(room) Then
                        foundFlag = True
                        panelTimeHandler(parsedStart, parsedEnd)
                    End If
                End If
            End While

            If Not foundFlag And Not genFlag Then
                If Not ignoreRoomMap.ContainsValue(room) Then
                    unfoundRoomMap.Add(room, room)
                End If
            End If

        Catch ex As Exception
            If Not missedScheduleMap.ContainsValue(devIDStr) Then
                missedScheduleMap.Add(panel, devIDStr)
                If Not lookupArray.Contains(s) Then
                    lookupArray.Add(s)
                End If
            End If
        End Try
    End Sub

    '
    'Method Name: panelTimeHandler()
    'Purpose: to avoid calling the insert function for start and end times that 
    '         are already occupied for that schedule
    'Accepts: two doubles containing the start and end times for a schedule
    '
    Public Sub panelTimeHandler(ByVal startTime As String, ByVal endTime As String)

        Dim temps As PanelData = New PanelData()

        ' insert flag used so that schedules don't get inserted a second time when executing the for loop
        Dim insertFlag As Boolean = False

        If Not firstRun Then
            temps.dev = devIDStr
            temps.inst = instanceStr
            temps.tempStart = startTime
            temps.tempEnd = endTime
            insert()
            firstRun = True

        Else
            Dim pd As PanelData = Nothing

            For index As Integer = 0 To tempPanelInfoArray.Count Step 1
                pd = tempPanelInfoArray.Item(index)
                If pd.dev.Equals(devIDStr) And pd.inst.Equals(instanceStr) Then
                    insertFlag = True
                    If pd.tempStart <= startTime And pd.tempEnd >= endTime Then
                        'Do nothing
                        'Break/ Exit If
                    Else
                        temps.dev = devIDStr
                        temps.inst = instanceStr
                        temps.tempStart = startTime
                        temps.tempEnd = endTime
                        insert()
                        tempPanelInfoArray.Item(index) = temps
                    End If
                End If
            Next

            If Not insertFlag Then
                temps.dev = devIDStr
                temps.inst = instanceStr
                temps.tempStart = startTime
                temps.tempEnd = endTime
                tempPanelInfoArray.Add(temps)
                insert()
            End If
        End If
    End Sub

    '
    'Method Name: convertTo24HoursFormat()
    'Purpose: to convert 12 hour base format to 24 hour format
    'Accepts: a string containing time in 12 hour format (3:00PM)
    'Returns: a date format of the passed in string changed to 24 hour format (15:00)
    '

    Public Shared Function convertTo24HoursFormat(ByVal twelveHourFormat As DateTime) As String
        'Dim time As DateTime = DateTime.ParseExact(twelveHourFormat, "hh: mma", Nothing)
        Dim convertedTime As String = Format(twelveHourFormat, "HH:mm:ss tt")
        Return convertedTime
        'Private TWELVE_TF As String = "hh:mma"
        'Private TWENTY_FOUR_TF As String = "HH:mm:ss"

    End Function

    '
    'Method Name: checkDesc()
    'Purpose: checks the description of to see whether or not it is "DO NOT BOOK" etc.
    'Accepts: a string containing the description of that schedule time
    'Returns: a boolean depending on the description comparison
    '
    Public Function checkDesc(ByVal d As String) As Boolean
        If d.Contains("DO NOT BOOK") OrElse d.Contains("HOLD DO NOT BOOK") OrElse d.Contains("HOLD - DO NOT BOOK") OrElse d.Contains("HOLD - DO NOT BOOK SSC ALCOVES") OrElse d.Contains("Do Not Book - UC") Then
            Return False
        Else
            Return True
        End If
    End Function

    '
    'Method Name: insert()
    'Purpose: Used to insert new times into the database for a specified schedule or 
    '         update current times in the database
    '
    ' Database Work
    '
    Private Sub insert()
        Dim updateSuccess As Integer = 0

        cmdChange = New ADODB.Command
        cmdChange.ActiveConnection = Cnxn

        Dim strtComm, endComm As String
        strtComm = finalStart.Substring(0, 8)
        endComm = finalEnd.Substring(0, 8)

        cmdStr = "update ARRAY_BAC_SCH_Schedule set SCHEDULE_TIME = {t '" & endComm & "'} WHERE SCHEDULE_TIME >=  {t '" &
                                strtComm & "'} AND" & " SCHEDULE_TIME <=  {t '" & endComm & "'} AND VALUE_ENUM = 0 AND DEV_ID = " &
                                devIDStr & " and INSTANCE = " & instanceStr & " and day = " & day
        cmdChange.CommandText = cmdStr
        Result = cmdChange.Execute()

        updateSuccess = Result.State

        If updateSuccess < 1 Then
            Dim found As Boolean = False

            cmdStr = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) " &
                    " values (2, " & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & strtComm & "'}, 1, 'Enum')"
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute()

            cmdStr = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & strtComm & "'} AND" &
            " VALUE_ENUM = 1 AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute()

            found = Result.State

            If Not found Then
                If invalidFlag = False Then

                    Dim msg As String = "Incorrect values were inserted for Start Time row, Device: " & devIDStr & " Instance: " & instanceStr & vbNewLine &
                           "Would you like to ignore these errors? "
                    Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
                         MsgBoxStyle.Critical
                    Dim response = Nothing

                    If response = MsgBox(msg, style) = MsgBoxResult.Yes Then
                        invalidFlag = True
                      End If

                    If lookupArray.Contains(panelLine) Then
                        lookupArray.Add(panelLine)
                    End If
                End If
            End If

            cmdStr = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) " &
                    " values (2," & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & endComm & "'}, NULL, 'NULL')"
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute

            cmdStr = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & endComm & "'} AND" &
            " DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute
            found = Result.State

            If Not found Then
                If invalidFlag = False Then

                    Dim msg As String = "Incorrect values were inserted for End Time row, Device: " & devIDStr & " Instance: " & instanceStr & vbNewLine &
                           "Would you like to ignore these errors? "
                    Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
                         MsgBoxStyle.Critical
                    Dim response = Nothing

                    If response = MsgBox(msg, style) = MsgBoxResult.Yes Then
                        invalidFlag = True
                    End If

                    If lookupArray.Contains(panelLine) Then
                        lookupArray.Add(panelLine)
                    End If
                End If
            End If
        End If
    End Sub

    '
    'Method Name: clearSchedule()
    'Purpose: Clears reservation schedules for same day as current schedules 
    '         prior to writing in the new schedule information 
    '
    ' Progress bar work and Database May need to rework matcher and for loop
    '
    Public Sub clearSchedule(ByVal s As String)

        Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(s)
        Dim strLine As String = fileReader.ReadLine()
        Dim firstFlag As Boolean = False 'flag used to make sure the variables get assigned properly in the for loop

        'While loop to go through text until its got panel info 
        While Not strLine.Contains("###")
            strLine = fileReader.ReadLine()
        End While

        While strLine <> Nothing And Not fileReader.EndOfStream
            Dim currentPanelInfoArray As String()
            currentPanelInfoArray = strLine.Split("#")

            If strLine.Contains("##") Then

                devIDStr = currentPanelInfoArray.ElementAt(6)
                instanceStr = currentPanelInfoArray.ElementAt(9)

                cmdChange = New ADODB.Command
                cmdChange.ActiveConnection = Cnxn

                cmdStr = "Delete FROM  ARRAY_BAC_SCH_Schedule WHERE INSTANCE = " & instanceStr & " And DEV_ID = " & devIDStr & " And DAY = " & day

                cmdChange.CommandText = cmdStr
                cmdChange.Execute()

                progress += 1
                progBar.Value = progress
                progBar.Refresh()

            End If
            strLine = fileReader.ReadLine()

        End While
    End Sub

    '
    'Method Name: writeNewLookup()
    'Purpose: Writes a new look-up table containing only the panels that were not 
    '         responding the first run through
    '
    Public Sub writeNewLookup()
        Try
            For index As Integer = 0 To lookupArray.Count Step 1
                My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\newLookup.txt", lookupArray.Item(index) & vbNewLine, True)
            Next
            lookupArray.Clear()

        Catch ex As Exception
            MsgBox("ERROR in writing new look-up: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: writeErrorLog()
    'Purpose: Writes the panel information for the panels that were not responding 
    '         also writes a list of the rooms that were not found in the look up table
    '
    Public Sub writeErrorLog()
        Try
            If Not missedScheduleMap.Count = 0 Then 'Dictionary isnt empty, missing some panels schedules
                Dim str As String = "Please query the following Panels, and then click the Generate Missed Schedules Button."
                TextBoxErrorLog.AppendText(vbNewLine & str & vbNewLine)
                For Index As Integer = 1 To missedScheduleMap.Last.Key Step 1
                    TextBoxErrorLog.AppendText("Panel: " & missedScheduleMap.ElementAt(Index).Key & ' Values and Key might need to be changed
                                               "Device Number: " & missedScheduleMap.ElementAt(Index).Value & vbNewLine)
                Next
            End If

            If Not unfoundRoomMap.Count = 0 Then 'Dictionary isnt empty, missing a few rooms
                TextBoxErrorLog.AppendText("Could not find Panel Information for the following rooms: " & vbNewLine)
                For Index As Integer = 1 To unfoundRoomMap.Last.Key Step 1
                    TextBoxErrorLog.AppendText("Room: " & unfoundRoomMap.ElementAt(Index).Value & vbNewLine) ' May need to change value
                Next
            End If
        Catch ex As Exception
            MsgBox("ERROR in writing error log: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: writeErrorLogFile()
    'Purpose: Writes the context of the error log to a text file
    '
    Public Sub writeErrorLogFile()
        Try
            My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\errorLog.txt", TextBoxErrorLog.Text, False)
            MsgBox("Error Log saved to W:\Reserve\RsvGeneratorFiles\errorLog.txt")

        Catch ex As Exception
            MsgBox("ERROR in writing Error Log File: " & ex.Message)
        End Try
    End Sub

    'Button Handlers

    'Browse Look-up Table
    Private Sub btnBrowseLookupTbl_Click(sender As Object, e As EventArgs) Handles btnBrowseLookupTbl.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.InitialDirectory = "W:\Reserve"
        fd.Filter = "txt files (*.txt)|*.txt|rec files (*.rec)|*.rec|All files (*.*)|*.*"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.Equals(vbOK) Then
            TextBoxBrowseLookup.Text = fd.FileName
        End If

        TextBoxBrowseLookup.Text = fd.FileName
    End Sub

    'Browse Schedule
    Private Sub btnBrowseSch_Click(sender As Object, e As EventArgs) Handles btnBrowseSch.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.InitialDirectory = "W:\Reserve\Daily"
        fd.Filter = "txt files (*.txt)|*.txt|rec files (*.rec)|*.rec|All files (*.*)|*.*"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.Equals(vbOK) Then
            TextBoxBrowseSch.Text = fd.FileName
        End If

        TextBoxBrowseSch.Text = fd.FileName
    End Sub

    'Generate Missed Schedules
    Private Sub btnGenMissedSch_Click(sender As Object, e As EventArgs) Handles btnGenMissedSch.Click
        genFlag = True
        progBar.Value = 0
        progress = 0
        'tempPanelInfoArray.Clear()
        Try
            Dim scheduleFile As String = TextBoxBrowseLookup.Text
            Dim newSch As String = TextBoxBrowseSch.Text ' Change These vars

            fileSize = scheduleFile.Length() + newSch.Length()
            progBar.Maximum = fileSize
            functionWorker()

        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try

    End Sub

    'Generate Schedules
    Private Sub btnGenSch_Click(sender As Object, e As EventArgs) Handles btnGenSch.Click
        firstRun = False
        progBar.Value = 0
        progress = 0

        tempPanelInfoArray.Clear()

        Try
            Dim scheduleFile As String = TextBoxBrowseSch.Text
            Dim newSch As String = TextBoxBrowseLookup.Text ' Change These vars
            fileSize = scheduleFile.Length() + newSch.Length()

            progBar.Maximum = fileSize

            'Call FunctionWorker
            functionWorker()
            btnGenSch.Enabled = True

        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    'Load All Notes
    Private Sub btnLoadAllNotes_Click(sender As Object, e As EventArgs) Handles btnLoadAllNotes.Click
        loadNotes()
    End Sub

    'Save note changes
    Private Sub btnSaveNoteChanges_Click(sender As Object, e As EventArgs) Handles btnSaveNoteChanges.Click
        saveNotes()
    End Sub

    'clear notes
    Private Sub btnClearNotes_Click(sender As Object, e As EventArgs) Handles btnClearNotes.Click
        clearNotes()
    End Sub

    'Submit notes
    Private Sub btnSubmitNotes_Click(sender As Object, e As EventArgs) Handles btnSubmitNotes.Click
        writeNotes()
    End Sub

    'Save error log
    Private Sub btnSaveErrorLog_Click(sender As Object, e As EventArgs) Handles btnSaveErrorLog.Click
        Try
            writeErrorLogFile()
        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: writeNotes()
    'Purpose: Writes the text from the notes TextBox in the program 
    '         into a text file for the user to view later
    '
    Public Sub writeNotes()
        Try
            My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\notes.txt", TextBoxNotePad.Text, False) ' Should re-write whole text file
            MsgBox("Notes submitted to W:\Reserve\RsvGeneratorFiles\notes.txt")

        Catch ex As Exception
            MsgBox("ERROR in writing notes: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: clearNotes()
    'Purpose: Clears the entire notes.txt file so that it is blank
    '
    Public Sub clearNotes()
        Try
            ' Create msg and style for a message box with options yes/no and a critical message icon.
            Dim msg As String = "Are you sure you want to delete all previous notes? "
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
                         MsgBoxStyle.Critical

            Dim response = MsgBox(msg, style)
            If response = MsgBoxResult.Yes Then
                'They want to clear all text in the notes.txt
                My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\notes.txt", "", False) ' Clear the .txt File

                TextBoxNotePad.Text = "Schedule Date: " & getDate() & vbNewLine & "Notes: "

                My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\notes.txt", TextBoxNotePad.Text, False) ' Recreate the notes.txt as a defualt
                MsgBox("Notes deleted from W:\Reserve\RsvGeneratorFiles\notes.txt")
            End If

        Catch ex As Exception
            MsgBox("ERROR in clearing notes: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: loadNotes()
    'Purpose: Loads the notes.txt file into the TextBox making it 
    '         available to edit 
    '
    Public Sub loadNotes()
        Try
            Dim fileReader As System.IO.StreamReader
            fileReader = My.Computer.FileSystem.OpenTextFileReader("W:\Reserve\RsvGeneratorFiles\notes.txt")
            TextBoxNotePad.Text = "" 'Make the current TextBox blank

            Dim strLine As String = ""
            strLine = fileReader.ReadLine()
            TextBoxNotePad.AppendText(strLine & vbNewLine)

            While Not fileReader.EndOfStream
                strLine = fileReader.ReadLine()
                TextBoxNotePad.AppendText(strLine & vbNewLine)
            End While
            fileReader.Close()
        Catch ex As Exception
            MsgBox("ERROR in loading notes: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: saveNotes()
    'Purpose: Overwrites the notes.txt file with the current text 
    '         in the TextBox
    '
    Public Sub saveNotes()
        Try
            'Create msg and style for a message box with options yes/no and a critical message icon.
            Dim msg As String = "Are you sure you want to save changes? "
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
                         MsgBoxStyle.Critical
            Dim response = MsgBox(msg, style)

            If response = MsgBoxResult.Yes Then
                'They want to save the note changes to the notes.txt
                My.Computer.FileSystem.WriteAllText("W:\Reserve\RsvGeneratorFiles\notes.txt", TextBoxNotePad.Text, False)
                MsgBox("Notes saved to W:\Reserve\RsvGeneratorFiles\notes.txt")
            End If

        Catch ex As Exception
            MsgBox("ERROR in saving notes: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: getDate()
    'Purpose: reads the schedule file to determine the date those schedules are for 
    'Returns: a string containing the date of the schedules from the selected schedule file
    '
    Public Function getDate() As String
        Dim d As String = ""
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(TextBoxBrowseSch.Text)

        Dim strLine As String = ""
        strLine = fileReader.ReadLine()
        'While strLine <> Nothing
        d = strLine.Substring(30, 55).Trim()
        'End While
        fileReader.Close()
        Return d
    End Function

    '
    'Method Name: defualtSettings()
    'Purpose: to reset the textboxes and buttons
    '
    Public Sub defaultSettings()
        TextBoxNotePad.Text = "Schedule Date: " & getDate() & vbNewLine & "Notes:"
        TextBoxErrorLog.Text = "Error Log: "
        btnGenMissedSch.Enabled = False
        btnGenSch.Enabled = True
    End Sub
End Class
