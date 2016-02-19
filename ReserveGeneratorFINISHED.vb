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
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Form1
    Public RevGeneratorVB()

    Private genFlag As Boolean = False
    Private ignoreFlag As Boolean = False
    Private firstRun As Boolean = False
    Private invalidFlag As Boolean = False
    Private updateFlag As Boolean = False

    Private TWELVE_TF As String = "hh:mma"
    Private TWENTY_FOUR_TF As String = "{HH:mm:ss}"
    Private finalEnd, finalStart, room, desc, instanceStr, devIDStr, panel As String

    'Database Variables
    Dim Cnxn As New ADODB.Connection
    Dim cmdChange As ADODB.Command
    Dim Result As ADODB.Recordset
    Dim cmdStr As String
    Dim strtComm, endComm, insert1, insert2, select1, select2, updateCall As String
    Dim version As String = ""

    Private day As Integer
    Private progress As Double = 0.0
    Private fileSize As Long
    Private panelLine As String = ""

    Dim roomMap As New Dictionary(Of String, String)
    Dim missedScheduleDictionairy As New Dictionary(Of String, String)
    Dim unfoundRoomDictionairy As New Dictionary(Of String, String)
    Dim ignoreRoomMap As New Dictionary(Of String, String)

    'ArrayLists
    Private tempPanelInfoArray As New ArrayList() 'Type PanelData
    Private lookupArray As New ArrayList()        'Type String

    'PanelData Class
    'Holds data related to the panels like the device Id, instance, start time and end time
    Public Class PanelData
        Public dev As String = ""
        Public inst As String = ""
        Public tempStart As Double = 0.0
        Public tempEnd As Double = 0.0
    End Class

    '
    'Method Name: Form1_Load()
    'Purpose: At run time of the application this method is auto called and will give 
    '         it the Date and make sure the generate missed schedules button Is disabled
    '
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoxNotePad.Text = "Schedule Date: " & getDate()
        TextBoxErrorLog.Text = "Notes: "
        btnGenMissedSch.Enabled = False
    End Sub

    '
    'Method Name: functionWorker()
    'Purpose: To run all the methods for generating the schedules
    '
    Public Sub functionWorker()
        Try
            defaultSettings()
            dbConnect()
            readFile(TextBoxBrowseSch.Text)
            dbClose()
            btnGenMissedSch.Enabled = True
            progBar.Value = progBar.Maximum
            writeErrorLog()
            writeNewLookup()
            btnGenSch.Enabled = False

            If missedScheduleDictionairy.Count = 0 And unfoundRoomDictionairy.Count = 0 Then
                MsgBox("Schedules Created!")
            Else
                MsgBox("Schedules Created with some errors." & vbNewLine & "Please veiw the Error Log below for further detail")
            End If

            missedScheduleDictionairy.Clear()

        Catch ex As Exception
            MsgBox("ERROR in generating schedule: " & ex.Message)
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
            TextBoxErrorLog.AppendText("Status: Connecting to Database... " & vbNewLine)
            Cnxn = New ADODB.Connection
            Cnxn.ConnectionString = "Dsn=Delta Network;uid=DELTA"
            Cnxn.Open()
            TextBoxErrorLog.AppendText("Status: Database Connected." & vbNewLine)

        Catch ex As SqlException
            System.Console.WriteLine("ERROR in opening database: " & ex.Message)
        End Try
    End Sub ' end dbConnect

    '
    ' Method Name: dbClose() 
    ' Purpose: Used to close the database connection to prevent a hanging connection
    '
    Private Sub dbClose()
        TextBoxErrorLog.AppendText("Status: Closing database connection..." & vbNewLine)
        Try
            Cnxn.Close()
            Cnxn = Nothing
            TextBoxErrorLog.AppendText("Status: Database Closed." & vbNewLine)

        Catch ex As SqlException
            System.Console.WriteLine("ERROR in closing database: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: readFile()
    'Purpose: reads the lines of a given file and performs regular expression 
    '         search for the necessary information
    '
    Public Sub readFile(ByRef filePath As String)
        Dim tempDay As Integer = 0
        Dim pattern As String = "((Monday)|(Tuesday)|(Wednesday)|(Thursday)|(Friday)|(Saturday)|(Sunday))"
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(filePath)

        Dim strLine As String = fileReader.ReadLine()
        Dim dayStr = strLine.Substring(0, 50).Trim()
        dayStr = dayStr.Substring(0, 9)

        'For loop to grab the day of the week
        If dayStr.Contains(",") Then
            For ind As Integer = 0 To dayStr.Length()
                If dayStr.ElementAt(ind) = "," Then
                    dayStr = dayStr.Substring(0, ind)
                    Exit For
                End If
            Next
        End If

        While Not fileReader.EndOfStream
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

            Dim values As String
            Dim startTime, endTime

            'Format string to get the time frame by itself
            strLine = strLine.Trim()
            values = strLine.Substring(56, 85).Trim()
            startTime = values.Substring(0, 7).Trim()

            'Convert to a DateTime object and -1 hour from it to account for warm-up on schedule
            Dim startTimeTemp As DateTime = startTime
            startTimeTemp = startTimeTemp.AddHours(-1)
            startTime = startTimeTemp
            endTime = values.Substring(9, 16).Trim()

            'Final time in 24 hour format
            finalStart = convertTo24HoursFormat(startTime)
            finalEnd = convertTo24HoursFormat(endTime)

            'Format string for the room value
            values = values.Substring(17).Trim()
            room = values.Substring(0, 25).Trim()

            'Format string for the description
            values = values.Substring(25).Trim()
            desc = values.Trim()

            If Not checkDesc(desc) Then
                'do not use this room. no need to allocate a schedule
            Else

                If day <> tempDay Then
                    TextBoxErrorLog.AppendText("Status: Clearing Schedules... " & vbNewLine)
                    If Not btnGenMissedSch.Enabled Then
                        clearSchedule(TextBoxBrowseLookup.Text)
                    ElseIf btnGenMissedSch.Enabled Then
                        clearSchedule(TextBoxGenMissedSch.Text)
                    End If
                    tempDay = day
                    TextBoxErrorLog.AppendText("Status: Creating new schedules..." & vbNewLine)
                End If

                If Not btnGenMissedSch.Enabled() Then
                    readLookupTable(TextBoxBrowseLookup.Text)
                ElseIf btnGenMissedSch.Enabled() Then
                    readLookupTable(TextBoxGenMissedSch.Text)
                End If

                progress += 1
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
        Dim strLine As String = ""
        Try
            ' found flag is used so that if the room for that schedule is not found in
            ' the Lookup table it will be put to the map to be displayed later
            Dim foundFlag As Boolean = False
            Dim skipFlag = False 'Check if room should be ignored
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(s)
            strLine = fileReader.ReadLine()

            While Not fileReader.EndOfStream
                If ignoreFlag = False Then
                    Dim lines As String = strLine.Substring(47)
                    Dim ignoreRooms() As String = lines.Split(", ")

                    For index As Integer = 0 To ignoreRooms.Length() - 1
                        ignoreRoomMap.Add(ignoreRooms.ElementAt(index), ignoreRooms.ElementAt(index))
                    Next

                    ignoreFlag = True

                ElseIf ignoreFlag = True Then

                    For i As Integer = 0 To ignoreRoomMap.Count() - 1
                        If ignoreRoomMap.ElementAt(i).Value.Trim() = room.Trim() Then
                            skipFlag = True
                            Exit While
                        End If
                    Next

                    Dim currentPanelInfoArray As String()
                    While Not strLine.Contains("###")
                        If fileReader.EndOfStream Then
                            Exit While
                        End If
                        strLine = fileReader.ReadLine()
                    End While

                    panelLine = strLine
                    roomMap.Clear()
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

                    Dim parsedStart As Double = Double.Parse(finalStart.Substring(0, 5).Replace(":", "."))
                    Dim parsedEnd As Double = Double.Parse(finalEnd.Substring(0, 5).Replace(":", "."))

                    For i As Integer = 0 To roomMap.Count - 1
                        If roomMap.ElementAt(i).Value.Trim() = room.Trim() And skipFlag = False Then
                            foundFlag = True
                            panelTimeHandler(parsedStart, parsedEnd)
                            Exit For
                        End If
                    Next
                End If

                strLine = fileReader.ReadLine()
            End While

            If Not foundFlag And Not genFlag And Not skipFlag Then
                'all flags are off so its a missed room
                unfoundRoomDictionairy.Add(room, room)
            End If

        Catch ex As Exception
            If Not missedScheduleDictionairy.ContainsValue(devIDStr) Then
                missedScheduleDictionairy.Add(panel, devIDStr)
                If Not lookupArray.Contains(s) Then
                    lookupArray.Add(strLine)
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
            updateFlag = False
            tempPanelInfoArray.Add(temps)
            insert()
            firstRun = True
        Else
            Dim pd As PanelData = Nothing
            For index As Integer = 0 To tempPanelInfoArray.Count - 1
                pd = tempPanelInfoArray.Item(index)
                If pd.dev.Equals(devIDStr) And pd.inst.Equals(instanceStr) Then
                    insertFlag = True
                    If pd.tempStart <= startTime And pd.tempEnd >= endTime Then

                    Else ' panel exists change values for it
                        temps.dev = devIDStr
                        temps.inst = instanceStr
                        temps.tempStart = startTime
                        temps.tempEnd = endTime

                        If startTime >= pd.tempStart And startTime <= pd.tempEnd And endTime <= pd.tempEnd Then 'Do nothing, no need to update or insert for a time slot already alocated
                            Exit For
                        ElseIf startTime >= pd.tempStart And startTime <= pd.tempEnd And endTime > pd.tempEnd Then 'only need to update no insert needed
                            updateFlag = True
                        ElseIf startTime > pd.tempStart And startTime > pd.tempEnd And endTime > pd.tempEnd Then 'new insert needed no update
                            updateFlag = False
                        Else
                            updateFlag = False
                            MsgBox("Else block hit in panelTimeHandler: check values?")
                        End If

                        insert()
                        tempPanelInfoArray.Item(index) = temps
                    End If
                End If
            Next

            'Add new panel
            If Not insertFlag Then
                updateFlag = False
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
    'Method Name: insert()
    'Purpose: Used to insert new times into the database for a specified 
    '         schedule Or update current times In the database
    '
    Private Sub insert()
        Dim found As Boolean = False

        cmdChange = New ADODB.Command
        cmdChange.ActiveConnection = Cnxn

        strtComm = finalStart.Substring(0, 8)
        endComm = finalEnd.Substring(0, 8)

        cmdStr = "SELECT ApplicationSWVersion FROM OBJECT_BAC_DEV WHERE DEV_ID = " & devIDStr
        cmdChange.CommandText = cmdStr

        version = cmdChange.Execute().GetString
        version = version.Trim()

        If version.Equals("V3.40") Then
            insert1 = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) values (1, " & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & strtComm & "'}, 1, 'Enum')"
            select1 = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & strtComm & "'} AND VALUE_ENUM = 1 AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            insert2 = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) values (1," & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & endComm & "'}, NULL, 'NULL')"
            select2 = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & endComm & "'} AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            updateCall = "Update ARRAY_BAC_SCH_Schedule Set SCHEDULE_TIME = {t '" & endComm & "'} WHERE SCHEDULE_TIME >=  {t '" & strtComm & "'} AND SCHEDULE_TIME <=  {t '" & endComm & "'} AND VALUE_TYPE = 'NULL' AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and day = " & day

        ElseIf version.Equals("V3.33") Or version.Equals("3.33") Then
            insert1 = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) values (1, " & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & strtComm & "'}, 1, 'Unsupported')"
            select1 = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & strtComm & "'} AND VALUE_ENUM = 1 AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            insert2 = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) values (1," & devIDStr & ", " & instanceStr & ", " & day & ", {t '" & endComm & "'}, 0, 'Unsupported')"
            select2 = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" & endComm & "'} AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and DAY = " & day
            updateCall = "Update ARRAY_BAC_SCH_Schedule Set SCHEDULE_TIME = {t '" & endComm & "'} WHERE SCHEDULE_TIME >=  {t '" & strtComm & "'} AND SCHEDULE_TIME <=  {t '" & endComm & "'} AND VALUE_ENUM = 0 AND DEV_ID = " & devIDStr & " and INSTANCE = " & instanceStr & " and day = " & day

        Else
            insert1 = ""
            select1 = ""
            insert2 = ""
            select2 = ""
            updateCall = ""
            MsgBox("Error: framework for panel is not 3.33 or 3.40. Program will not be able to process this panel.")
        End If

        If updateFlag = True Then
            cmdStr = updateCall
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute()
        Else
            cmdStr = insert1
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute()

            cmdStr = select1
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute()
            found = Result.State

            If Not found Then
                If invalidFlag = False Then
                    Dim msg As String = "Incorrect values were inserted for Start Time row, Device: " & devIDStr & " Instance: " & instanceStr & vbNewLine & "Would you like to ignore these errors? "
                    Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
                    Dim response = Nothing

                    If response = MsgBox(msg, style) = MsgBoxResult.Yes Then
                        invalidFlag = True
                    End If

                    If lookupArray.Contains(panelLine) Then
                        lookupArray.Add(panelLine)
                    End If
                End If
            End If

            cmdStr = insert2
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute

            cmdStr = select2
            cmdChange.CommandText = cmdStr
            Result = cmdChange.Execute
            found = Result.State

            If Not found Then
                If invalidFlag = False Then
                    Dim msg As String = "Incorrect values were inserted for End Time row, Device: " & devIDStr & " Instance: " & instanceStr & vbNewLine & "Would you like to ignore these errors? "
                    Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
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
    Public Sub clearSchedule(ByVal s As String)
        Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(s)
        Dim strLine As String = fileReader.ReadLine()

        If fileReader.EndOfStream = True Then
            Exit Sub
        End If

        'flag used to make sure the variables get assigned properly in the for loop
        Dim firstFlag As Boolean = False

        While Not strLine.Contains("###")
            strLine = fileReader.ReadLine()
        End While

        While Not fileReader.EndOfStream
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
    'Purpose: Writes a new look-up table containing only the panels 
    '         that were Not responding the first run through
    '
    Public Sub writeNewLookup()
        Try
            My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\newLookup.txt", "", False)
            If lookupArray.Count <> 0 Then
                For index As Integer = 0 To lookupArray.Count - 1
                    My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\newLookup.txt", lookupArray.Item(index) & vbNewLine, True) 'W:\WES SOFTWARE\Reservations Generator\Test Schedule Online\newLookup.txt
                Next
                lookupArray.Clear()
            End If
        Catch ex As Exception
            MsgBox("ERROR in writing new look-up: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: writeErrorLog()
    'Purpose: Writes the panel information for the panels that were Not 
    '         responding also writes a list of the rooms that were Not 
    '         found In the look uptable
    '
    Public Sub writeErrorLog()
        Try
            If Not missedScheduleDictionairy.Count = 0 Then 'Dictionary isnt empty, missing some panels schedules
                Dim str As String = "Please query the following Panels, and then click the Generate Missed Schedules Button."
                TextBoxErrorLog.AppendText(vbNewLine & str & vbNewLine)
                For Index As Integer = 0 To missedScheduleDictionairy.Count - 1
                    TextBoxErrorLog.AppendText("Panel: " & missedScheduleDictionairy.ElementAt(Index).Key & "   Device Number: " & missedScheduleDictionairy.ElementAt(Index).Value & vbNewLine)
                Next
            End If

            If Not unfoundRoomDictionairy.Count = 0 Then 'Dictionary isnt empty, missing a few rooms
                TextBoxErrorLog.AppendText(vbNewLine & "Could not find Panel Information for the following rooms: " & vbNewLine)
                For Index As Integer = 0 To unfoundRoomDictionairy.Count - 1
                    TextBoxErrorLog.AppendText("Room: " & unfoundRoomDictionairy.ElementAt(Index).Value & vbNewLine)
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
            My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\errorLog.txt", TextBoxErrorLog.Text, False)
            MsgBox("Error Log saved to W:\WES SOFTWARE\Reservations Generator\errorLog.txt")
        Catch ex As Exception
            MsgBox("ERROR in writing Error Log File: " & ex.Message)
        End Try
    End Sub

    'Button Handlers

    'Browse Look-up Table
    Private Sub btnBrowseLookupTbl_Click(sender As Object, e As EventArgs) Handles btnBrowseLookupTbl.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.InitialDirectory = "W:\WES SOFTWARE\Reservations Generator"
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
        fd.InitialDirectory = "W:\WES SOFTWARE\Reservations Generator"
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
        defaultSettings()
        genFlag = True
        progBar.Value = 0
        progress = 0
        tempPanelInfoArray.Clear()
        Try
            setProgressBarMax()
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
            setProgressBarMax()
            functionWorker()
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
            My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\notes.txt", TextBoxNotePad.Text, False)
            MsgBox("Notes submitted to W:\WES SOFTWARE\Reservations Generator\notes.txt")
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
            Dim msg As String = "Are you sure you want to delete all previous notes? "
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
                         MsgBoxStyle.Critical
            Dim response = MsgBox(msg, style)

            If response = MsgBoxResult.Yes Then
                My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\notes.txt", "", False)
                TextBoxNotePad.Text = "Schedule Date: " & getDate() & vbNewLine & "Notes: "
                My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\notes.txt", TextBoxNotePad.Text, False)
                MsgBox("Notes deleted from W:\WES SOFTWARE\Reservations Generator\notes.txt")
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
            fileReader = My.Computer.FileSystem.OpenTextFileReader("W:\WES SOFTWARE\Reservations Generator\notes.txt")
            TextBoxNotePad.Text = ""
            Dim strLine As String = fileReader.ReadLine()
            TextBoxNotePad.AppendText(strLine & vbNewLine)

            While Not fileReader.EndOfStream
                strLine = fileReader.ReadLine()
                TextBoxNotePad.AppendText(strLine & vbNewLine)
            End While
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
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
            Dim response = MsgBox(msg, style)

            If response = MsgBoxResult.Yes Then
                My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\notes.txt", TextBoxNotePad.Text, False)
                MsgBox("Notes saved to W:\WES SOFTWARE\Reservations Generator\notes.txt")
            End If
        Catch ex As Exception
            MsgBox("ERROR in saving notes: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: getDate()
    'Purpose: Reads the schedule file to determine the date those 
    '         schedules are For 
    'Returns: A string containing the date of the schedules from 
    '         the selected schedule file
    '
    Public Function getDate() As String
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(TextBoxBrowseSch.Text)
        Dim strLine As String = fileReader.ReadLine().Trim()
        strLine = strLine.Substring(0, 55).Trim()
        Return strLine
    End Function

    '
    'Method Name: defualtSettings()
    'Purpose: to reset the textboxes and buttons to default values
    '
    Public Sub defaultSettings()
        TextBoxNotePad.Text = "Schedule Date: " & getDate() & vbNewLine & "Notes:"
        TextBoxErrorLog.Text = "Error log: " & vbNewLine
    End Sub

    '
    'Method Name: setProgressBarMax()
    'Purpose: To go through each file and get a line count to 
    '         use as a maximum value for the progress bar
    '
    Public Sub setProgressBarMax()
        Dim count As Integer = 0
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(TextBoxBrowseSch.Text)
        Dim readLn As String = fileReader.ReadLine()

        While Not fileReader.EndOfStream
            count += 1
            readLn = fileReader.ReadLine()
        End While

        fileReader = My.Computer.FileSystem.OpenTextFileReader(TextBoxBrowseLookup.Text)
        readLn = fileReader.ReadLine()

        While Not fileReader.EndOfStream
            count += 1
            readLn = fileReader.ReadLine()
        End While

        progBar.Maximum = count + 19

    End Sub

    '
    'Method Name: convertTo24HoursFormat()
    'Purpose: to convert 12 hour base format to 24 hour format
    'Accepts: a string containing time in 12 hour format (3:00PM)
    'Returns: a date format of the passed in string changed to 24 hour format (15:00)
    '
    Public Shared Function convertTo24HoursFormat(ByVal twelveHourFormat As DateTime) As String
        Dim convertedTime As String = Format(twelveHourFormat, "HH:mm:ss tt")
        Return convertedTime
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
End Class
