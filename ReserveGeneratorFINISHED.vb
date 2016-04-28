'
'Program Name: RsvGeneratorVB
'Purpose: Reads in all the Panels and corresponding information from a text file.
'         Reads in the Schedule information and loads them to proper variables from a text file.
'         Writes the schedule to the database.
'         Creates an error log containing panels that were not responding and rooms that do not exist on the system
'         Creates an updated version of the panel look up table containing just the panels that were not responding the first time through.
'         
'Date Started: Janurary 13th, 2016 
'Date Finished: Feburary 29th 2016 Updated April 1st and 19th
'Programmer and Coder: Eric Odette
'
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Form1
    Public RevGeneratorVB()

    'Boolean Flags
    Private genFlag As Boolean = False          ' Flag to determine if the programs already run once
    Private ignoreFlag As Boolean = False       ' Makes sure the program only runs the block of code once
    Private firstRun As Boolean = False         ' Makes sure the panel object hasnt been created already
    Private invalidFlag As Boolean = False      ' Monitors the inserts into the database dont fail
    Private updateFlag As Boolean = False       ' Monitors whether or not the the update is needed
    Private errorFlag As Boolean = False        ' Determines what message will be outputted at the end

    'Time and Panel variables
    Private finalEnd, finalStart, room, desc, instanceStr, devIDStr, panel As String

    'Database Variables
    Private Cnxn As New ADODB.Connection
    Private cmdChange As ADODB.Command
    Private Result As ADODB.Recordset
    Private strtComm, endComm, insert1, insert2, select1, select2, updateCall As String
    Private version As String = ""
    Private cmdStr As String = ""

    'File and panel variables
    Private day As Integer = 0
    Private progress As Double = 0.0
    Private fileSize As Long = 0
    Private panelLine As String = ""

    'Dictionary variables 
    Private roomMap As New Dictionary(Of String, String)
    Private missedScheduleDictionairy As New Dictionary(Of String, String)
    Private unfoundRoomDictionairy As New Dictionary(Of String, String)
    Private ignoreRoomMap As New Dictionary(Of String, String)

    'ArrayLists
    Private tempPanelInfoArray As New ArrayList() 'Type PanelData
    Private lookupArray As New ArrayList()        'Type String

    'PanelData Class
    'Holds data related to the panels like the device Id, instance, start time and end time
    Public Class PanelData
        Private dev As String = ""
        Private inst As String = ""
        Private tempStart As Double = 0.0
        Private tempEnd As Double = 0.0

        'Property functions (Act as getters and setters for VB.NET)			
        Public Property Dev1 As String
            Get
                Return dev
            End Get
            Set(value As String)
                dev = value
            End Set
        End Property

        Public Property Inst1 As String
            Get
                Return inst
            End Get
            Set(value As String)
                inst = value
            End Set
        End Property

        Public Property TempStart1 As Double
            Get
                Return tempStart
            End Get
            Set(value As Double)
                tempStart = value
            End Set
        End Property

        Public Property TempEnd1 As Double
            Get
                Return tempEnd
            End Get
            Set(value As Double)
                tempEnd = value
            End Set
        End Property
    End Class

    '
    'Method Name: Form1_Load()
    'Purpose: At run time of the application this method is auto called and will give 
    '         it the Date and make sure the generate missed schedules button Is disabled
    '
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dateLbl.Text = "Schedule Date: " & getDate()
        dateLbl.Refresh()
        TextBoxErrorLog.Text = "Note Log: "
        btnGenMissedSch.Enabled = False
        errorFlag = False
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

            If missedScheduleDictionairy.Count = 0 And unfoundRoomDictionairy.Count = 0 And errorFlag = False Then
                MsgBox("Schedules Created!")
            ElseIf errorFlag = False Then
                MsgBox("Schedules Created with some errors." & vbNewLine & "Please veiw the Error Log below for further detail")

            End If

            missedScheduleDictionairy.Clear()

        Catch ex As Exception
            MsgBox("ERROR in FunctionWorker. Message: " & ex.Message)
            progBar.Value = 0
            If Cnxn.Equals(True) Then
                dbClose()
            End If
            errorFlag = True
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
            MsgBox("ERROR in opening the database connection. Message: " & ex.Message & vbNewLine & "Check that the database connection is opertaing properly")
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
            MsgBox("ERROR in closing the database connection. Message: " & ex.Message & vbNewLine & "Check that the database connection is opertaing properly")
        End Try
    End Sub

    '
    'Method Name: readFile()
    'Purpose: reads the lines of a given file and performs regular expression 
    '         search for the necessary information
    '
    Public Sub readFile(ByRef filePath As String)
        Dim dayStr As String = ""
        Try
            Dim tempDay As Integer = 0
            Dim pattern As String = "((Monday)|(Tuesday)|(Wednesday)|(Thursday)|(Friday)|(Saturday)|(Sunday))"
            Dim fileReader As System.IO.StreamReader
            fileReader = My.Computer.FileSystem.OpenTextFileReader(filePath)

            Dim strLine As String = fileReader.ReadLine()
            dayStr = strLine.Substring(0, 50).Trim()
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
            fileReader = My.Computer.FileSystem.OpenTextFileReader(filePath)
            While Not fileReader.EndOfStream And strLine <> ""
                strLine = fileReader.ReadLine()
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
                values = strLine.Substring(54, 85).Trim()
                startTime = values.Substring(0, 7).Trim()
                endTime = values.Substring(9, 16).Trim()

                'Format string for the room value
                values = values.Substring(17).Trim()
                room = values.Substring(0, 25).Trim()

                'Format string for the description
                values = values.Substring(25).Trim()
                desc = values.Trim()

                'Convert to a DateTime object and -1 hour from it to account for warm-up on schedule
                Try
                    Dim startTimeTemp As DateTime = startTime
                    startTimeTemp = startTimeTemp.AddHours(-1)
                    startTime = startTimeTemp
                Catch ex As Exception
                    TextBoxErrorLog.AppendText("ERROR converting start time to a DateTime variable. start time = """ & startTime & """. Room: " & room & vbNewLine & "Please ensure that the value is in a proper time format (00:00:00, 00:00, etc.)" & vbNewLine)
                    errorFlag = True
                    Exit Sub
                End Try

                'Final time in 24 hour format
                finalStart = convertTo24HoursFormat(startTime)
                finalEnd = convertTo24HoursFormat(endTime)

                If checkDesc(desc) Then
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
            End While
        Catch ex As Exception
            TextBoxErrorLog.AppendText("ERROR in reading the schedule file. Message from error: " & ex.Message & vbNewLine &
                   " Data when it failed " & vbNewLine & " Panel: " & devIDStr & "Time frame: " & finalStart & " - " & finalEnd & "Day: " & dayStr)
            errorFlag = True
            Exit Sub
        End Try
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
            temps.Dev1 = devIDStr
            temps.Inst1 = instanceStr
            temps.TempStart1 = startTime
            temps.TempEnd1 = endTime
            updateFlag = False
            tempPanelInfoArray.Add(temps)
            insert()
            firstRun = True
        Else
            Dim pd As PanelData = Nothing
            For index As Integer = 0 To tempPanelInfoArray.Count - 1
                pd = tempPanelInfoArray.Item(index)
                If pd.Dev1.Equals(devIDStr) And pd.Inst1.Equals(instanceStr) Then
                    insertFlag = True
                    If pd.TempStart1 <= startTime And pd.TempEnd1 >= endTime Then

                    Else ' panel exists change values for it
                        temps.Dev1 = devIDStr
                        temps.Inst1 = instanceStr
                        temps.TempStart1 = startTime
                        temps.TempEnd1 = endTime

                        If startTime >= pd.TempStart1 And startTime <= pd.TempEnd1 And endTime <= pd.TempEnd1 Then 'Do nothing, no need to update or insert for a time slot already alocated
                            Exit For
                        ElseIf startTime >= pd.TempStart1 And startTime <= pd.tempEnd1 And endTime > pd.tempEnd1 Then 'only need to update no insert needed
                            updateFlag = True
                        ElseIf startTime > pd.tempStart1 And startTime > pd.tempEnd1 And endTime > pd.tempEnd1 Then 'new insert needed no update
                            updateFlag = False
                        Else
                            updateFlag = False
                            TextBoxErrorLog.AppendText("Else block hit in panelTimeHandler. Panel: " & temps.Dev1 & ". Time Frame: " & temps.TempStart1 & " To " & temps.TempEnd1 & vbNewLine _
                                   & "Check Panel to make sure its online or that the file is getting proper time values")
                            errorFlag = True
                            Exit Sub
                        End If

                        insert()
                        tempPanelInfoArray.Item(index) = temps
                    End If
                End If
            Next

            'Add new panel
            If Not insertFlag Then
                updateFlag = False
                temps.Dev1 = devIDStr
                temps.Inst1 = instanceStr
                temps.TempStart1 = startTime
                temps.TempEnd1 = endTime
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

        version = cmdChange.Execute().GetString()
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
            MsgBox("Error: framework for panel is not 3.33 or 3.40. Program will not be able to process this panel. Panel type: " & version & "Panel: " & devIDStr)
            Exit Sub
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
        Try
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
        Catch ex As Exception
            MsgBox("ERROR in clearing schedules: Failed on line '" & strLine & "' in the look-up table. " & vbNewLine & "Please check that panel to see if it is offline, or causing issues.")
        End Try
    End Sub

    '
    'Method Name: writeNewLookup()
    'Purpose: Writes a new look-up table containing only the panels 
    '         that were Not responding the first run through
    '
    Public Sub writeNewLookup()
        Try
            'Dim fileWriter As StreamWriter = File.CreateText("W:\WES SOFTWARE\Reservations Generator\newLookup.txt")
            'fileWriter.Write("", False)
            My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\newLookup.txt", "", False)

            If lookupArray.Count <> 0 Then
                For index As Integer = 0 To lookupArray.Count - 1
                    My.Computer.FileSystem.WriteAllText("W:\WES SOFTWARE\Reservations Generator\newLookup.txt", lookupArray.Item(index) & vbNewLine, True)
                Next
                lookupArray.Clear()
            End If
            'fileWriter.Close()
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
            MsgBox("ERROR in writing error log to the program, Message: " & ex.Message)
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
            MsgBox("ERROR in writing Error Log File, Message: " & ex.Message)
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
        btnGenSch.Enabled = True
        btnGenMissedSch.Enabled = False
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
        dateLbl.Text = "Schedule Date: " & getDate()
        dateLbl.Refresh()
        btnGenSch.Enabled = True
        btnGenMissedSch.Enabled = False
    End Sub

    'Generate Missed Schedules
    Private Sub btnGenMissedSch_Click(sender As Object, e As EventArgs) Handles btnGenMissedSch.Click
        defaultSettings()
        genFlag = True
        progBar.Value = 0
        progress = 0
        tempPanelInfoArray.Clear()
        setProgressBarMax()
        functionWorker()
    End Sub

    'Generate Schedules
    Private Sub btnGenSch_Click(sender As Object, e As EventArgs) Handles btnGenSch.Click
        defaultSettings()
        firstRun = False
        progBar.Value = 0
        progress = 0
        tempPanelInfoArray.Clear()
        setProgressBarMax()
        functionWorker()
    End Sub

    'Save error log
    Private Sub btnSaveErrorLog_Click(sender As Object, e As EventArgs) Handles btnSaveErrorLog.Click
        Try
            writeErrorLogFile()
        Catch ex As Exception
            System.Console.WriteLine("ERROR saving the Error Log, Message: " & ex.Message)
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
        fileReader.Close()
        Return strLine
    End Function

    '
    'Method Name: defualtSettings()
    'Purpose: to reset the textboxes and buttons to default values
    '
    Public Sub defaultSettings()
        dateLbl.Text = "Schedule Date: " & getDate()
        dateLbl.Refresh()
        TextBoxErrorLog.Text = "Note log: " & vbNewLine
        TextBoxErrorLog.Refresh()
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
