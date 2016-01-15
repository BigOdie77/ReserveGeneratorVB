/*
 * Program Name: RsvGenerator
 * Purpose: Reads in all the Panels an corresponding information from a text file.
 * 			Reads in the schedule information and loads it to variables from a text file.
 * 			Writes the schedules to the database.
 * 			Creates an error log containing panels that were not responding and rooms that do not exist on the system.
 * 			Creates an updated version of the panel look up table containing just the panels what were not responding the first time through.
 * 			Creates a note file containing notes submitted from the user when creating a schedule.
 * Date: 14/09/2011
 * Author: Justin A. Janda
 */

import java.awt.*;
import java.awt.event.*;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Writer;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.*;
import javax.swing.filechooser.FileFilter;

public class RsvGenerator extends JFrame
{
	private static final long serialVersionUID = 1L;
	private static final double FR_WIDTH_FACTOR = 0.5;
	private static final double FR_HEIGHT_FACTOR = 0.55;
	
	private boolean genFlag = false;
	private boolean ignoreFlag = false;
	private boolean firstRun = false;
	private boolean invalidFlag = false;
	
	private static final DateFormat TWELVE_TF = new SimpleDateFormat("hh:mma");
	private static final DateFormat TWENTY_FOUR_TF = new SimpleDateFormat("HH:mm:ss");
	
	private String finalEnd, finalStart, room, desc, instanceStr, devIDStr, panel;
	private JButton browseButton, genButton, lookupButton, noteButton, clearNoteButton, loadNoteButton, saveButton, rerunButton, writeLogButton;
	private JTextField scheduleTxt, lookupTxt, rerunTxt;
	private JTextArea noteText, errorText;
	private JProgressBar progressBar;
	
	// DB connection and variables
	private Connection conn = null;
	private Statement stmt = null;
	private ResultSet rslt = null;
	
	private int day;
	private double progress = 0.0;
	private long fileSize;

	private String panelLine = "";
	
	private HashMap<String, String> roomMap = new HashMap<String, String>();
	private HashMap<String, String> missedScheduleMap = new HashMap<String, String>();
	private HashMap<String, String> unfoundRoomMap = new HashMap<String, String>();
	
	private ArrayList<PanelData> panelTempInfoArray = new ArrayList<PanelData>(1);
	private ArrayList<String> lookupArray = new ArrayList<String>();
	
	public class PanelData {
		String dev = "";
		String inst = "";
		double tempStart = 0.0;
		double tempEnd = 0.0;
	}
	
	/* 
	 * Creates a SwingWorker thread and runs all the main functions in a background thread
	 * so the progress bar can be updated displaying the current progress of the tasks
	 */
	private class FunctionWorker extends SwingWorker<Void, Void> {
		@Override
		protected void done() {
			super.done();
		}

		@Override
		protected Void doInBackground() throws Exception {	
			try
			{
				genButton.setEnabled(false);
				dbConnect();
				readFile(scheduleTxt.getText());
				dbClose();
				genButton.setEnabled(true);
				progressBar.setValue(progressBar.getMaximum());
				errorText.setText("Error Log: ");
				writeErrorLog();
				writeNewLookup();
				
				if(missedScheduleMap.isEmpty() &&  unfoundRoomMap.isEmpty())
					JOptionPane.showMessageDialog(null, "Schedules Created!");
				else if(!unfoundRoomMap.isEmpty() && missedScheduleMap.isEmpty())
					JOptionPane.showMessageDialog(null, "Schedules Created, with some errors. \nPlease view Error Log below for further details.");
				else
				{
					JOptionPane.showMessageDialog(null, "Schedules Created, with some errors. \nPlease view Error Log below for further details.");
					rerunButton.setEnabled(true);
				}
				missedScheduleMap.clear();
			}
			catch (Exception e) {
				JOptionPane.showMessageDialog(null, "doInBackGround - Error: " + e.getMessage());
				errorText.setText("Error Log: ");
				progressBar.setValue(0);
				dbClose();
			}
			return null;
		}
	}
	
	private class TextFilter extends FileFilter {
	
		public boolean accept(File f) {
			if (f.isDirectory())
				return true;
			String s = f.getName();
			int i = s.lastIndexOf('.');
			
			if (i > 0 && i < s.length() - 1)
				
			if (s.substring(i + 1).toLowerCase().equals("txt") || s.substring(i + 1).toLowerCase().equals("rec"))
				return true;
			
			return false;
		}
		
		public String getDescription() {
			return "Accepts txt and rec files only.";
		}
	}

	public RsvGenerator() throws HeadlessException, IOException {

		setTitle("Reservation Schedule Generator");
		setSize( (int)(getToolkit().getScreenSize().width * FR_WIDTH_FACTOR), 
					(int)(getToolkit().getScreenSize().height * FR_HEIGHT_FACTOR));
		setLocationRelativeTo(null); 
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		
		setLayout( new BorderLayout());
		
		JPanel northPanel = new JPanel();
		northPanel.setLayout( new GridLayout(4,2,10,10));
		buttonListener btnListener = new buttonListener();
				
		// creates the northPanel and all of its components
		scheduleTxt = new JTextField("W:\\Reserve\\Daily\\dailyevents.rec");
		scheduleTxt.setEditable(false);
		browseButton = new JButton("Browse for Schedule");
		browseButton.addActionListener(btnListener);
		scheduleTxt.scrollRectToVisible(getBounds());
			
		lookupTxt = new JTextField("W:\\Reserve\\RsvGeneratorFiles\\panelLookupTable.txt");
		lookupTxt.setEditable(false);
		lookupButton = new JButton("Browse for Look up Table");
		lookupButton.addActionListener(btnListener);
		lookupTxt.scrollRectToVisible(getBounds());
		
		rerunTxt = new JTextField("W:\\Reserve\\RsvGeneratorFiles\\newLookup.txt");
		rerunTxt.setEditable(false);
		rerunButton = new JButton("Generate Missed Schedules");
		rerunButton.setEnabled(false);
		rerunButton.addActionListener(btnListener);
		rerunTxt.scrollRectToVisible(getBounds());
		
		UIManager.put("ProgressBar.selectionBackground", Color.BLACK);
		UIManager.put("ProgressBar.selectionForeground", Color.BLACK);
		
		progressBar = new JProgressBar();
		progressBar.setValue(0);
		progressBar.setStringPainted(true);
		genButton = new JButton("Generate Schedules");
		genButton.setEnabled(true);
		genButton.addActionListener(btnListener);
		
		northPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		northPanel.add(lookupTxt);
		northPanel.add(lookupButton);
		northPanel.add(scheduleTxt);
		northPanel.add(browseButton);
		northPanel.add(rerunTxt);
		northPanel.add(rerunButton);
		northPanel.add(progressBar);
		northPanel.add(genButton);
		
		add(northPanel, BorderLayout.NORTH);
		
		// Creates the centerPanel and all of its components
		JPanel centerPanel = new JPanel();
		centerPanel.setLayout( new GridLayout(2,1,10,10));
		centerPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		
		noteText = new JTextArea("Schedule Date: " + getDate() + "\nNotes: ");
		noteText.setLineWrap(true);
		noteText.setWrapStyleWord(true);
		noteText.setFont(new Font("Times New Roman", Font.BOLD, 14));
		JScrollPane scrollPane = new JScrollPane(noteText);
		centerPanel.add(scrollPane);
		
		errorText = new JTextArea("Error Log: ");
		errorText.setLineWrap(true);
		errorText.setWrapStyleWord(true);
		errorText.setEditable(false);
		errorText.setFont(new Font("Times New Roman", Font.BOLD, 14));
		JScrollPane scrollPane1 = new JScrollPane(errorText);
		centerPanel.add(scrollPane1);
		
		add(centerPanel, BorderLayout.CENTER);
		
		// Creates the southPanel and all of its components
		JPanel southPanel = new JPanel();
		southPanel.setLayout( new GridLayout(1,0,10,10));
		southPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		noteButton = new JButton("Submit Notes");
		noteButton.addActionListener(btnListener);
		clearNoteButton = new JButton("Clear All Notes");
		clearNoteButton.addActionListener(btnListener);
		loadNoteButton = new JButton("Load All Notes");
		loadNoteButton.addActionListener(btnListener);
		saveButton = new JButton("Save Note Changes");
		saveButton.addActionListener(btnListener);
		writeLogButton = new JButton("Save Error Log");
		writeLogButton.addActionListener(btnListener);
		
		southPanel.add(loadNoteButton);
		southPanel.add(saveButton);
		southPanel.add(clearNoteButton);
		southPanel.add(noteButton);
		southPanel.add(writeLogButton);
		add(southPanel, BorderLayout.SOUTH);		
		
		northPanel.setBackground(new Color(102, 0, 153));
		centerPanel.setBackground(new Color(102, 0, 153));
		southPanel.setBackground(new Color(102, 0, 153));
		progressBar.setForeground(Color.GREEN);
		
		addWindowListener(new WindowListener()
		{
			@Override
			public void windowActivated(WindowEvent arg0){}
			@Override
			public void windowClosed(WindowEvent arg0){}
			@Override
			public void windowClosing(WindowEvent arg0)
			{
				dbClose();
				RsvGenerator.this.dispose();
			}
			@Override
			public void windowDeactivated(WindowEvent arg0){}
			@Override
			public void windowDeiconified(WindowEvent arg0){}
			@Override
			public void windowIconified(WindowEvent arg0){}
			@Override
			public void windowOpened(WindowEvent arg0){}
		});
		setVisible(true);
	}

	public static void main(String[] args) throws HeadlessException, IOException
	{
		new RsvGenerator();
	}
	
	private class buttonListener implements ActionListener
	{
		public void actionPerformed(ActionEvent e)
		{
			if( e.getSource() == genButton )
			{				
				firstRun = false;
				progressBar.setIndeterminate(false);
				progressBar.setValue(0);
				progress = 0;
				panelTempInfoArray.clear();
				try 
				{
					File file = new File(scheduleTxt.getText());
					File file2 = new File(lookupTxt.getText());
					fileSize = file.length() + file2.length();
					progressBar.setMaximum((int) (fileSize));
					
					FunctionWorker task = new FunctionWorker();
					task.execute();
					genButton.setEnabled(true);
				} 
				catch (Exception e1){
					System.err.println("Error: " + e1.getMessage());
				}
			}
			if( e.getSource() == rerunButton )
			{
				genFlag = true;
				progressBar.setIndeterminate(false);
				progressBar.setValue(0);
				progress = 0;
				panelTempInfoArray.clear();
				try 
				{
					File file = new File(scheduleTxt.getText());
					File file2 = new File(rerunTxt.getText());
					fileSize = file.length() + file2.length();
					progressBar.setMaximum((int) (fileSize));
					
					FunctionWorker task = new FunctionWorker();
					task.execute();	
					
				} 
				catch (Exception e1){
					System.err.println("Error: " + e1.getMessage());
				}	
			}
			if( e.getSource() == browseButton )
			{
				final JFileChooser fc = new JFileChooser();
				fc.setCurrentDirectory(new File("W:\\Reserve\\Daily"));
				fc.setAcceptAllFileFilterUsed(false);
				fc.addChoosableFileFilter(new TextFilter());
				int returnVal = fc.showOpenDialog(RsvGenerator.this);

		        if (returnVal == JFileChooser.APPROVE_OPTION) 
		        {
		            File file = fc.getSelectedFile();
		            scheduleTxt.setText(file.getAbsolutePath());
		            try 
		            {
						noteText.setText("Schedule Date: " + getDate() + "\nNotes: ");
					} 
		            catch (IOException e1) 
		            {
						e1.printStackTrace();
					}
		        } 
			}
			if( e.getSource() == lookupButton )
			{
				final JFileChooser fc = new JFileChooser();
				fc.setAcceptAllFileFilterUsed(false);
				fc.setCurrentDirectory(new File("W:\\Reserve"));
				fc.addChoosableFileFilter(new TextFilter());
				int returnVal = fc.showOpenDialog(RsvGenerator.this);

		        if (returnVal == JFileChooser.APPROVE_OPTION) 
		        {
		            File file = fc.getSelectedFile();
		            lookupTxt.setText(file.getAbsolutePath());
		        } 
			}
			if( e.getSource() == noteButton )
			{
				writeNotes();
			}
			if( e.getSource() == clearNoteButton )
			{
				clearNotes();
			}
			if( e.getSource() == loadNoteButton )
			{
				loadNotes();
			}
			if( e.getSource() == saveButton )
			{
				saveNotes();	
			}
			if( e.getSource() == writeLogButton )
			{
				try 
				{
					writeErrorLogFile();
				} catch (IOException e1) 
				{
					e1.printStackTrace();
				}	
			}
		}
	}
		
	/*
	 * Method Name: readFile
	 * Purpose: reads the lines of a given file and
	 * performs regular expression search the necessary information
	 */
	public void readFile (String filePath) throws IOException, SQLException, ParseException
	{
			int tempDay = 0;
			FileInputStream fstream = new FileInputStream(filePath);
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			
			String strLine;
			String regex =  "((Monday)|(Tuesday)|(Wednesday)|(Thursday)|(Friday)|(Saturday)|(Sunday))"; 
		    Pattern pattern = Pattern.compile(regex);
			while ((strLine = br.readLine()) != null)   
			{
				Matcher matcher = pattern.matcher(strLine);
		        boolean matchFound = matcher.find();
		     
		        if (matchFound) {	
		        	if(matcher.group(1).equals("Monday"))
		        		day = 1;
		        	else if(matcher.group(1).equals("Tuesday"))
		        		day = 2;
		        	else if(matcher.group(1).equals("Wednesday"))
		        		day = 3;
		        	else if(matcher.group(1).equals("Thursday"))
		        		day = 4;
		        	else if(matcher.group(1).equals("Friday"))
		        		day = 5;
		        	else if(matcher.group(1).equals("Saturday"))
		        		day = 6;
		        	else if(matcher.group(1).equals("Sunday"))
		        		day = 7;
		        }
		        
				if(day != tempDay)
				{
					errorText.setText("Status: Clearing schedules...");
					if(!rerunButton.isEnabled())
						clearSchedules(lookupTxt.getText());
					else if (rerunButton.isEnabled())
						clearSchedules(rerunTxt.getText());
					tempDay = day;
				}

					String time = strLine.substring(85, 110).trim();
					String startTime, endTime;
					
					startTime = time.substring(0, 7).trim();
					endTime = time.substring(9).trim();
					
					finalStart = convertTo24hoursFormat(startTime);
					finalEnd = convertTo24hoursFormat(endTime);
					
					Date date = new Date();
					Calendar c = Calendar.getInstance();
					date = TWENTY_FOUR_TF.parse(finalStart);
					c.setTime(date);
					c.add(Calendar.HOUR, -1);
					finalStart = TWENTY_FOUR_TF.format(c.getTime());
					
					room = strLine.substring(112, 141).trim();
					desc = strLine.substring(138, 190).trim();

	
					if(!checkDesc(desc)) {}
	
					else
					{	
						errorText.setText("Status: Creating new schedules...");
						if(!rerunButton.isEnabled())
							readLookupTable(lookupTxt.getText());
						else if (rerunButton.isEnabled())
							readLookupTable(rerunTxt.getText());
						
						progress = progress + (strLine.getBytes().length);
						progressBar.setValue((int) progress);	
					}
			}
			in.close();
	}
	
	/*
	 * Method Name: readLookupTable
	 * Purpose: performs a regular expression search on the look up table to get panel information
	 */
	public void readLookupTable (String s) throws IOException
	{	
		String strLine = "";
		try
		{
			FileInputStream fstream = new FileInputStream(s);	
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));

			String regex =  "(([^#])*)(#{3})+" +
							"(([^#])*)(#{3})+" +
							"(([^#])*)(#{3})+" +
							"(([^#])*)"; 
			
		    Pattern pattern = Pattern.compile(regex);
		    
		    // foundFlag is used so that if the room for that schedule is
		    // not found in the look up table it will write it to the map
		    // to be displayed later
		    boolean foundFlag = false;
		    
		    HashMap<String, String> ignoreRoomMap = new HashMap<String, String>();
			while ((strLine = br.readLine()) != null)   
			{
				if(ignoreFlag == false)
				{	
					String line = strLine.substring(47);
					String[] ignoreRooms = line.split(", ");
		            for(int i = 0; i < ignoreRooms.length; ++i){
		            	ignoreRoomMap.put(ignoreRooms[i], ignoreRooms[i]);
		            }
		            ignoreFlag = true;
				}
				else if(ignoreFlag == true)
				{
					panelLine = strLine;
					roomMap.clear();
					 Matcher matcher = pattern.matcher(strLine);
				        boolean matchFound = matcher.find();
				        if (matchFound) {
				            panel = matcher.group(1);
				            String roomList = matcher.group(4);
				            devIDStr = matcher.group(7);
				            instanceStr = matcher.group(10);
				            String[] rooms = roomList.split(", ");
				            for(int i = 0; i < rooms.length; ++i){
				            	roomMap.put(rooms[i], rooms[i]);
				            }
				        }
				        
				        double parsedStart = Double.parseDouble(finalStart.substring(0, 5).replace(':', '.'));
						double parsedEnd = Double.parseDouble(finalEnd.substring(0, 5).replace(':', '.'));
						if(roomMap.containsValue(room))
						{
				        	foundFlag = true;
							panelTimeHandler(parsedStart, parsedEnd);
						}
				}
			}
			if(!foundFlag && !genFlag) {
				if(!ignoreRoomMap.containsValue(room))
				{
				if(!unfoundRoomMap.containsValue(room))
					unfoundRoomMap.put(room, room);
				}
			}
					
		}
		catch (Exception e){	
			if(!missedScheduleMap.containsValue(devIDStr))
				missedScheduleMap.put(panel, devIDStr);
			if(!lookupArray.contains(strLine))
			{
				lookupArray.add(strLine);	
			}
		}
	}
	
	/*
	 * Method Name: panelTimeHandler
	 * Accepts: two doubles containing the start and end times
	 * Purpose: to avoid calling the insert function for start and end times
	 *  that are already occupied for that schedule in the database
	 */
	public void panelTimeHandler (double timeStart, double timeEnd) throws SQLException {	
		PanelData temps = new PanelData();

		// insertFlag is used so that schedules dont get inserted a second time when exiting the for loop
		boolean insertFlag = false;
		if(!firstRun)
		{
			temps.dev = devIDStr;
			temps.inst = instanceStr;
			temps.tempStart = timeStart;
			temps.tempEnd = timeEnd;
			panelTempInfoArray.add(temps);
			insert();
			firstRun = true;
		}

		else
		{
			PanelData pd = null;
			for(int i = 0; i < panelTempInfoArray.size(); i++) {
				pd = panelTempInfoArray.get(i);
				if(pd.dev.equals(devIDStr) && pd.inst.equals(instanceStr)) {
					insertFlag = true;
					if(pd.tempStart <= timeStart  && pd.tempEnd >= timeEnd ) {
						// Do Nothing
						break;
					}
					else 
					{
						temps.dev = devIDStr;
						temps.inst = instanceStr;
						temps.tempStart = timeStart;
						temps.tempEnd = timeEnd;
						insert();
						panelTempInfoArray.set(i, temps);
						break;
					}
				}
			}
			if (!insertFlag)
			{
				temps.dev = devIDStr;
				temps.inst = instanceStr;
				temps.tempStart = timeStart;
				temps.tempEnd = timeEnd;
				panelTempInfoArray.add(temps);
				insert();
			}
		}
	}
	
  /*
   * Method Name: convertTo24hoursFormat
   * Accepts: a string containing time in 12 hour format(3:00PM)
   * Returns: a date format of the passed in string in 24 hour format(15:00);
   * Purpose: to convert 12 hour format to 24 hour format
   */
  public static String convertTo24hoursFormat(String twelveHourTime)
	        throws ParseException {
	    return TWENTY_FOUR_TF.format(
	            TWELVE_TF.parse(twelveHourTime));
  }
  
  /*
   * Method Name: checkDesc
   * Accepts: a string containing the description of that schedule time
   * Returns: a boolean depending on the description comparison
   * Purpose: checks the description to see whether to book this time slot or not
   */
  public boolean checkDesc (String d)
  {
	  if(d.equals("DO NOT BOOK") || d.equals("HOLD - DO NOT BOOK") || d.equals("HOLD DO NOT BOOK") || d.equals("HOLD - DO NOT BOOK SSC ALCOVES"))
		  return false;
	  else
		  return true;
  }
	
  /*
   * Method Name: dbConnect
   * Purpose: to connect to the database using the Delta Network odbc data source
   */
  private void dbConnect() throws SQLException
	{
		try
		{
			errorText.setText("Status: Connecting to database...");
			Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
			conn = DriverManager.getConnection("jdbc:odbc:Delta Network");
			errorText.setText("Status: Database connected...");
		}
		catch(ClassNotFoundException ex)
		{
			System.out.println(ex.getMessage());
		}	
	}// end dbConnect
  
  /*
   * Method Name: dbClose
   * Purpose: Closes the connection to the database to prevent a hanging connection
   */
	private void dbClose()
	{
		errorText.setText("Status: Closing database connection...");
			try {
				if(rslt!=null)
				rslt.close();
			} 
			catch (SQLException e2) {
				e2.printStackTrace();
			}
			try {
				if(stmt!=null)
				stmt.close();
			} 
			catch (SQLException e1) {
				e1.printStackTrace();
			}
			try {
				if(conn!=null)
				conn.close();
			} 
			catch (SQLException e) {
				e.printStackTrace();
			}
	}// end dbClose
		
	/*
	 * Method Name: insert
	 * Purpose: Used to insert new times into the database for a specified schedule or update current times in the database
	 */
	private void insert() throws SQLException
	{
		stmt = conn.createStatement();	
		int updateSuccess = 0;
		String sqlStm = "update ARRAY_BAC_SCH_Schedule set SCHEDULE_TIME = {t '" + finalEnd + "'} WHERE SCHEDULE_TIME >=  {t '" + finalStart + "'} AND" +
			" SCHEDULE_TIME <=  {t '" + finalEnd + "'} AND VALUE_ENUM = 0 AND DEV_ID = " + devIDStr + " and INSTANCE = " + instanceStr + " and day = " + day;
		
		updateSuccess = stmt.executeUpdate(sqlStm);
		
		if (updateSuccess < 1)
		{	
			boolean found = false;
			sqlStm = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) " +
					" values (1, " + devIDStr + ", " + instanceStr + ", " + day + ", {t '" + finalStart + "'}, 1, 'Unsupported')";
			stmt.executeUpdate(sqlStm);
			
			sqlStm = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" + finalStart + "'} AND" +
			" VALUE_ENUM = 1 AND DEV_ID = " + devIDStr + " and INSTANCE = " + instanceStr + " and DAY = " + day;
			ResultSet rslt = stmt.executeQuery(sqlStm);
			found = rslt.next();
			if(!found)
			{
				if(invalidFlag == false)
				{
					int a = JOptionPane.showConfirmDialog(null, "Incorrect values were inserted for Start Time row, Device: " + devIDStr + " Instance: " + instanceStr + 
							"\nWould you like to ignore these errors? ", "Error", JOptionPane.YES_NO_OPTION);
					if( a == JOptionPane.YES_OPTION)
					{
						invalidFlag = true;
					}
				}
				if(!lookupArray.contains(panelLine))
				{
					lookupArray.add(panelLine);	
				}
			}
				
			
			sqlStm = "insert into ARRAY_BAC_SCH_Schedule (SITE_ID, DEV_ID, INSTANCE, DAY, SCHEDULE_TIME, VALUE_ENUM, Value_Type) " +
					" values (1," + devIDStr + ", " + instanceStr + ", " + day + ", {t '" + finalEnd + "'}, 0, 'Unsupported')";
			stmt.executeUpdate(sqlStm);
			
			sqlStm = "Select * from ARRAY_BAC_SCH_Schedule where SCHEDULE_TIME =  {t '" + finalEnd + "'} AND" +
			" VALUE_ENUM = 0 AND DEV_ID = " + devIDStr + " and INSTANCE = " + instanceStr + " and DAY = " + day;
			 rslt = stmt.executeQuery(sqlStm);
			 found = rslt.next();
			 if(!found)
			 {
				if(invalidFlag == false)
				{
					int a = JOptionPane.showConfirmDialog(null, "Incorrect values were inserted for End Time row, Device: " + devIDStr + " Instance: " + instanceStr + 
							"\nWould you like to ignore these errors? ", "Error", JOptionPane.YES_NO_OPTION);
					if( a == JOptionPane.YES_OPTION)
					{
						invalidFlag = true;
					}
				}
				if(!lookupArray.contains(panelLine))
				{
					lookupArray.add(panelLine);	
				}
			}
		}
		if(stmt!=null)
			stmt.close();
	}
	
	/*
	 * Method Name: clearSchedules
	 * Purpose: clears reservation schedules for the same day as current schedules
	 * prior to writing in the new schedule information
	 */
	public void clearSchedules (String s) throws IOException, SQLException
	{
			FileInputStream fstream = new FileInputStream(s);
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;

			String regex =  "(([^#])*)(#{3})+" +
							"(([^#])*)(#{3})+" +
							"(([^#])*)(#{3})+" +
							"(([^#])*)"; 
		    Pattern pattern = Pattern.compile(regex);
		    
			while ((strLine = br.readLine()) != null)   
			{
				 Matcher matcher = pattern.matcher(strLine);
			        boolean matchFound = matcher.find();
			        if (matchFound){
			            devIDStr = matcher.group(7);
			            instanceStr = matcher.group(10);

			            stmt = conn.createStatement();	
						
			    		String sqlStm = "Delete FROM  ARRAY_BAC_SCH_Schedule WHERE INSTANCE = " + instanceStr + " AND DEV_ID = " + devIDStr + " AND DAY = "  + day;
			    		stmt.executeUpdate(sqlStm);
			    		
			    		int bytes = (strLine.getBytes().length);
			    		progress = progress + bytes;
						progressBar.setValue((int) progress);
						progressBar.repaint();

			    		if(stmt!=null)
			    			stmt.close();  
		            }
	        }
	}
	
	/*
	 * Method Name: writeNewLookup
	 * Purpose: Writes a new look up table containing only the panels that were not responding the first run through
	 */
	public void    () throws IOException {
		try
		{
			Writer output = null;
			File file3 = new File("W:\\Reserve\\RsvGeneratorFiles\\newLookup.txt");
			output = new BufferedWriter(new FileWriter(file3));
			for(int i = 0; i < lookupArray.size(); ++i)
			{
				output.write(lookupArray.get(i) + System.getProperty("line.separator"));
			}
			output.close();
			lookupArray.clear();
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, " writeNewLookup - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: writeErrorLog
	 * Purpose: Writes the panel information for the panels that were not responding
	 * also writes a list of the rooms that were not found in the look up table
	 */
	public void writeErrorLog () throws IOException {
		try
		{
			if(!missedScheduleMap.isEmpty()) {
				String str = "Please query the following Panels, and then click Generate Missed Schedules.";
				errorText.append(System.getProperty("line.separator") + str + System.getProperty("line.separator"));
				for(@SuppressWarnings("rawtypes") Map.Entry entry : missedScheduleMap.entrySet()) {
					errorText.append("Panel: " + entry.getKey() + "    Device Number: " + entry.getValue() + System.getProperty("line.separator"));
				};
			}
			
			if(!unfoundRoomMap.isEmpty()) {
				errorText.append(System.getProperty("line.separator")+ "Could not find Panel Information for the following rooms: " + System.getProperty("line.separator"));
				for(@SuppressWarnings("rawtypes") Map.Entry entry : unfoundRoomMap.entrySet()) {
					errorText.append("Room: " + entry.getKey() + System.getProperty("line.separator"));
				};
			}
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, "writeErrorLog - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: writeErrorLogFile
	 * Purpose: writes the context of the error log to a text file.
	 */
	public void writeErrorLogFile () throws IOException {
		Writer output = null;
		File file4 = new File("W:\\Reserve\\RsvGeneratorFiles\\errorLog.txt");
		output = new BufferedWriter(new FileWriter(file4));
		String log = errorText.getText();
		output.write(log);
		output.close();
		JOptionPane.showMessageDialog(null, "Error Log saved to W:\\Reserve\\RsvGeneratorFiles\\errorLog.txt");
	}
	
	/*
	 * Method Name: writeNotes
	 * Purpose: Writes the the text from the notes fields in the program into a text file for the user to view later
	 */
	public void writeNotes () throws StringIndexOutOfBoundsException {
		try
		{
			Writer output = null;
			File file4 = new File("W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
			output = new BufferedWriter(new FileWriter(file4, true));	
			String line = noteText.getText();
			StringBuffer sb = new StringBuffer(line);
			output.write(sb.toString().replace("\n", System.getProperty("line.separator")));
			output.write(System.getProperty("line.separator"));
			output.write(System.getProperty("line.separator") + System.getProperty("line.separator"));
			output.close();
			JOptionPane.showMessageDialog(null, "Notes submitted to W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, " writeNotes - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: clearNotes
	 * Purpose: Clears the entire notes.txt file so that it is blank
	 */
	public void clearNotes () throws StringIndexOutOfBoundsException {
		try
		{
			int answer = JOptionPane.showConfirmDialog(null, "Are you sure you want to delete all previous notes?", "Warning", JOptionPane.YES_NO_OPTION);
			if( answer == JOptionPane.YES_OPTION)
			{
				Writer output = null;
				File file5 = new File("W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
				output = new BufferedWriter(new FileWriter(file5));	
				output.write("");
				output.close();
				noteText.setText("Schedule Date: " + getDate() + "\nNotes: ");
				JOptionPane.showMessageDialog(null, "Notes deleted from W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
			}
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, " clearNotes - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: loadNotes
	 * Purpose: Loads the notes.txt file into the JTextArea making them available to edit
	 */
	public void loadNotes () throws StringIndexOutOfBoundsException {
		try
		{
			FileInputStream fstream = new FileInputStream("W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			noteText.setText("");
			while ((strLine = br.readLine()) != null)   
			{
				noteText.append(strLine);
				noteText.append(System.getProperty("line.separator"));
			}
			in.close();
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, "loadNotes - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: saveNotes
	 * Purpose: Overwrites the notes.txt with the current text in the JTextArea
	 */
	public void saveNotes () throws StringIndexOutOfBoundsException {
		try
		{
			int answer = JOptionPane.showConfirmDialog(null, "Are you sure you want to save changes?", "Warning", JOptionPane.YES_NO_OPTION);
			if( answer == JOptionPane.YES_OPTION)
			{
				Writer output = null;
				File file5 = new File("W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
				output = new BufferedWriter(new FileWriter(file5));	
				output.write(noteText.getText());
				output.close();
				JOptionPane.showMessageDialog(null, "Notes Saved to W:\\Reserve\\RsvGeneratorFiles\\notes.txt");
			}
			
		}
		catch (Exception e){
			JOptionPane.showMessageDialog(null, " saveNotes - Error: " + e.getMessage());
		}
	}
	
	/*
	 * Method Name: getDate
	 * Purpose: reads the schedule file to determine the date those schedules are for
	 * Returns: a string containing the date of the schedules from the selected schedule file
	 */
	public String getDate () throws IOException
	{
			String d = "";
			FileInputStream fstream = new FileInputStream(scheduleTxt.getText());
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			
			String strLine;
			while ((strLine = br.readLine()) != null)   
			{
				d = strLine.substring(30, 60).trim();
			}
			in.close();
			return d;
	}
}
