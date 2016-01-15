Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Win32

Public Class RevGeneratorVB
    'Global Variables
    Private Const serialVersionUID As Long = 1L
    Private Const FR_WIDTH_FACTOR As Double = 0.5
    Private Const FR_HEIGHT_FACTOR As Double = 0.55

    'Boolean Flags
    Private genFlag As Boolean = False
    Private ignoreFlag As Boolean = False
    Private firstRun As Boolean = False
    Private invalidFlag As Boolean = False

    'Private Const TWELVE_TF As DateFormat = New DateFormat("h:mm tt")
    Private TWELVE_TF As String = "hh:mma"
    Private TWENTY_FOUR_TF As String = "HH:mm:ss"

    Private finalEnd, finalStart, room, desc, instanceStr, devIDStr, panel As String

    'DB Connection * MAY NEED TO BE EDITED *
    Dim connection As SqlConnection = New SqlConnection("Provider=MSDASQL.1; Password=LOGIN; USER ID=DELTA; Data Source= Delta Network;")
    Dim da As SqlDataAdapter
    Dim dr As DataRow
    Dim cmdBuilder As SqlCommandBuilder

    Private day As Integer
    Private progress As Double = 0.0
    Private fileSize As Long

    Private panelLine As String = ""

    'No hashtables in vb use Dictionary instead or something like it
    Dim roomMap As New Dictionary(Of String, String)
    Dim missedScheduleMap As New Dictionary(Of String, String)
    Dim unfoundRoomMap As New Dictionary(Of String, String)

    Private panelTempInfoArray As New ArrayList() 'Type PanelData
    Private lookupArray As New ArrayList()        'Type String

    'PanelData Class
    Public Class PanelData
        Dim dev As String = ""
        Dim inst As String = ""
        Dim tempStart As Double = 0.0
        Dim tempEnd As Double = 0.0
    End Class

    '
    'Method Name: textFilter()
    'Purpose: To make sure file extensions are in lower case and are either .txt or .rec
    '
	Public Sub textFilter()
	
	End Sub

    '
    'Method Name: Main()
    'Purpose: to execute the main chunks of code
    '
    Public Sub Main()
        'Execute program code? Or not

    End Sub

    'Button Handlers -----------------------------------------------------------------------------------------------------------------------------------------------
    'Browse for Look-up Table Button Handler
    Private Sub btnBrowseLookupTbl_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnBrowseLookupTbl.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.InitialDirectory = "W:\\Reserve"
        fd.Filter = "txt files (*.txt)|*.txt|rec files (*.rec)|*.rec|All files (*.*)|*.*"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult Then
            TextBoxBrowseLookup.Text = fd.FileName
        End If
    End Sub

    'Browse for Schedule Button Handler
    Private Sub btnBrowseSch_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnBrowseSch.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.InitialDirectory = "W:\\Reserve\\Daily"
        fd.Filter = "txt files (*.txt)|*.txt|rec files (*.rec)|*.rec|All files (*.*)|*.*"
        fd.FilterIndex = 3
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult Then
            TextBoxBrowseLookup.Text = fd.FileName
            Try
                TextBoxNotePad.AppendText("Schedule Date: " & getDate() & vbNewLine & "Notes:")
            Catch ex As Exception
                System.Console.WriteLine("Error: " & ex.Message)
            End Try

        End If
    End Sub

    'Generate Missed Schedules Button Handler
    Private Sub btnGenMissedSch_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnGenMissedSch.Click
        genFlag = True
        ProgBar.Value = 0
        progress = 0
        'panelTempInfoArray.Clear()
        Try
			Dim scheduleFile As String = TextBoxBrowseLookup
			Dim newSch As Sting = TextBoxBrowseSch ' Change These vars
			
			fileSize = scheduleFile.length() + newSch.length()
			progBar.setMaximum((int) fileSize)
			
			'Call FunctionWorker
			'Do the program
			
        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try

    End Sub

    'Generate Schedule Button Handler
    Private Sub btnGenSch_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnGenSch.Click
        firstRun = False
        ProgBar.Value = 0
        progress = 0

		panelTempInfoArray.Clear()
		
        Try
			Dim scheduleFile As String = TextBoxBrowseSch 
			Dim newSch As Sting = TextBoxBrowseLookup ' Change These vars
			
			fileSize = scheduleFile.length() + newSch.length()
			progBar.setMaximum((int) fileSize)
			
			'Call FunctionWorker

			genButton.IsEnabled = True
			
			
        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    'Load All Notes Button Handler
    Private Sub btnLoadAllNotes_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnLoadAllNotes.Click
        loadNotes()
    End Sub

    'Save Note Changes Button Handler
    Private Sub btnSaveNoteChanges_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSaveNoteChanges.Click
        saveNotes()
    End Sub

    'Clear Notes Button Handler
    Private Sub btnClearNotes_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnClearNotes.Click
        clearNotes()
    End Sub

    'Submit Notes Button Handler
    Private Sub btnSubmitNotes_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSubmitNotes.Click
        writeNotes()
    End Sub

    'Save Error Log Button Handler
    Private Sub btnSaveErrorLog_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSaveErrorLog.Click
        Try
            'writeErrorLog()
        Catch ex As Exception
            System.Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
	
    '
    'Method Name: readFile()
    'Purpose: reads the lines of a given file and performs regular expression search for the 
    '         necessary information
    '
	Public Sub readFile()
	
	tempDay = 0
	End Sub

    '
    'Method Name: readLookupTable()
    'Purpose: performs a regular expression search on the look up table to get panel information
    '
	Public Sub readLookupTable(s As String)
	
	End Sub

    '
    'Method Name: panelTimeHandler()
    'Purpose: to avoid calling the insert function for start and end times that 
    '         are already occupied for that schedule
    'Accepts: two doubles containing the start and end times for a schedule
    '
	Public panelTimeHandler(startTime As Double, endTime As Double)
	
	End Sub

    '
    'Method Name: convertTo24HoursFormat()
    'Purpose: to convert 12 hour base format to 24 hour format
    'Accepts: a string containing time in 12 hour format (3:00PM)
    'Returns: a date format of the passed in string changed to 24 hour format (15:00)
    '

    '                        - MAY NEED TO EDIT -

    Public Shared Function convertTo24HoursFormat(ByVal twelveHourFormat As String) As String
        Dim time As DateTime = DateTime.ParseExact(twelveHourFormat, "hh:mma", Nothing)
        Dim convertedTime As String = time.ToString("HH:mm:ss")
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
        If d.Contains("DO NOT BOOK") OrElse d.Contains("HOLD DO NOT BOOK") OrElse d.Contains("HOLD - DO NOT BOOK") OrElse d.Contains("HOLD - DO NOT BOOK SSC ALCOVES") Then
            Return False
        Else
            Return True
        End If
    End Function

    '
    'Method Name: dbConnect()
    'Purpose: Connects to the database using the Delta Network obdc data source
    '
    '*MAY NEED TO BE REVISED*

    Private Sub dbConnect()
        Try
            TextBoxErrorLog.Text = "Status: Connecting to Database... "
            'cn.ConnectionString = "Server=.;Database=northwind;UID=sa;PWD=;"
            'cn.Open()

            connection.Open() 'Database is now open
            TextBoxErrorLog.Text = "Status: Database connected..."

        Catch ex As SqlException
            System.Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Sub ' end dbConnect

    '
    ' Method Name: dbClose() 
    ' Purpose: Used to close the database connection to prevent a hanging connection
    '
    '*MAY NEED TO BE REVISED*

    Private Sub dbClose()

        TextBoxErrorLog.Text = "Status: Closing database connection..."
        Try
            If connection IsNot Nothing Then
                connection.Close()
            End If
        Catch ex As SqlException
            System.Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Sub

    '
    'Method Name: insert()
    'Purpose: Used to insert new times into the database for a specified schedule or 
    '         update current times in the database
    '
    Private Sub insert()
        Dim updateSuccess As Integer = 0
        Dim sqlStm As String = "update ARRAY_BAC_SCH_Schedule set SCHEDULE_TIME = {t '" & finalEnd & "'} WHERE SCHEDULE_TIME >=  {t '" & _
                                finalStart & "'} AND" & " SCHEDULE_TIME <=  {t '" & finalEnd & "'} AND VALUE_ENUM = 0 AND DEV_ID = " & _
                                devIDStr & " and INSTANCE = " & instanceStr & " and day = " & day
        'da = New SqlDataAdapter("select * from 



    End Sub

    '
    'Method Name: clearSchedule()
    'Purpose: Clears reservation schedules for same day as current schedules 
    '         prior to writing in the new schedule information 
    '
	Public Sub clearSchedule()
	
	End Sub

    '
    'Method Name: writeNewLookup()
    'Purpose: Writes a new look-up table containing only the panels that were not 
    '         responding the first run through
    '
	Public Sub writeNewLookup()
	
	End Sub

    '
    'Method Name: writeErrorLog()
    'Purpose: Writes the panel information for the panels that were not responding 
    '         also writes a list of the rooms that were not found in the look up table
    '
    Public Sub writeErrorLog()
        Try
            If Not missedScheduleMap.Count = 0 Then
                Dim str As String = "Please query the following Panels, and then click the Generate Missed Schedules Button."
                TextBoxErrorLog.AppendText(vbNewLine & str & vbNewLine)

                'want to go through the dictionary that holds the panels and room #'s that were not found and output them to the errorTextBox
                'For index As Integer missedScheduleMap. 

                'Next




            End If
        Catch ex As Exception

        End Try
    End Sub

    '
    'Method Name: writeErrorLogFile()
    'Purpose: Writes the context of the error log to a text file
    '
    Public Sub writeErrorLogFile()
        Dim errorLogFileLoc As String = "W:\\Reserve\\RsvGeneratorFiles\\errorLog.txt"

        Dim fileReader As String = My.Computer.FileSystem.ReadAllText(errorLogFileLoc)
        My.Computer.FileSystem.WriteAllText(errorLogFileLoc, TextBoxErrorLog.Text, True)
        MsgBox("Error Log saved to W:\\Reserve\\RsvGeneratorFiles\\errorLog.txt")

    End Sub

    '
    'Method Name: writeNotes()
    'Purpose: Writes the text from the notes TextBox in the program 
    '         into a text file for the user to view later
    '
    Public Sub writeNotes()
        Try
            'Writer output = null;
            Dim noteFileLoc As String = "W:\\Reserve\\RsvGeneratorFiles\\notes.txt"

            Dim fileReader As String = My.Computer.FileSystem.ReadAllText(noteFileLoc)
            'output = new BufferedWriter(new FileWriter(file4, true));	
            Dim line As String = TextBoxNotePad.Text()
            'StringBuffer sb = new StringBuffer(line);
            'output.write(sb.toString().replace("\n", System.getProperty("line.separator")));
            'output.write(System.getProperty("line.separator"));
            'output.write(System.getProperty("line.separator") + System.getProperty("line.separator"));
            'output.close();
            My.Computer.FileSystem.WriteAllText(noteFileLoc, TextBoxNotePad.Text, True)

            MsgBox("Notes submitted to W:\\Reserve\\RsvGeneratorFiles\\notes.txt")

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
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or _
                         MsgBoxStyle.Critical

            ' Display the message box
            MsgBox(msg, style)

            'save the response for an if statement to determine what each option will do
            Dim response = MsgBox(msg, style)

            If response = MsgBoxResult.Yes Then
                'They want to save the note changes to the notes.txt
                Dim noteFile As System.IO.StreamWriter
                noteFile = My.Computer.FileSystem.OpenTextFileWriter("W:\\Reserve\\RsvGeneratorFiles\\notes.txt", True)

                'May need to use a loop to get full box if it only does a line of text
                noteFile.WriteLine("")
                TextBoxNotePad.Text = "Schedule Date: " & getDate() & vbNewLine & "Notes: "
                MsgBox("Notes deleted from W:\\Reserve\\RsvGeneratorFiles\\notes.txt")
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
            Dim fileReader As String = ""
            fileReader = My.Computer.FileSystem.ReadAllText("W:\\Reserve\\RsvGeneratorFiles\\notes.txt")

            Dim strLine As String = ""
            TextBoxNotePad.Text = ""
            While fileReader IsNot Nothing
                strLine = fileReader
                TextBoxNotePad.AppendText(strLine)
                'TextBoxNotePad.AppendText(Uses System.getProperty("line.seperator"))
                'May not need, or find something to complete task
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
            ' Create msg and style for a message box with options yes/no and a critical message icon.
            Dim msg As String = "Are you sure you want to save changes? "
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or _
                         MsgBoxStyle.Critical
            'save the response for an if statement to determine what each option will do
            Dim response = MsgBox(msg, style)

            ' Display the message box
            MsgBox(msg, style)

            If response = MsgBoxResult.Yes Then
                'They want to save the note changes to the notes.txt
                Dim noteFile As System.IO.StreamWriter
                noteFile = My.Computer.FileSystem.OpenTextFileWriter("W:\\Reserve\\RsvGeneratorFiles\\notes.txt", True)

                'May need to use a loop to get full box if it only does a line of text
                noteFile.WriteLine(TextBoxNotePad.Text)

                MsgBox("Notes saved to W:\\Reserve\\RsvGeneratorFiles\\notes.txt")
            End If

        Catch ex As Exception
            'Console.WriteLine("ERROR in saving notes: " & ex.Message)
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

        Dim fileReader As String = ""
        fileReader = My.Computer.FileSystem.ReadAllText(TextBoxBrowseSch.Text)
        'Uses a DataInputStream and BuffedReader to get line by line? 
        '                       -MAY NOT WORK-
        Dim strLine As String

        While fileReader IsNot Nothing
            strLine = fileReader
            d = strLine.Substring(30, 60).Trim()
        End While

        Return d
    End Function
End Class
