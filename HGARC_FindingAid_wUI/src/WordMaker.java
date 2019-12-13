import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.*;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordMaker extends JPanel implements ActionListener
{
    static private final String newline = "\n";
    JButton fileButton, foldButton;
    static JTextArea log;
    JFileChooser fc;
    
	private static ArrayList<XWPFParagraph> content;
	private static XWPFDocument doc;
	private static XWPFParagraph activePara = null;
	private static XWPFRun activeRun = null;
	
	private static String version = "WordMaker v1.3.2";

    public WordMaker() 
    {
        super(new BorderLayout());

        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        
        fileButton = new JButton("Convert a File");
        fileButton.addActionListener(this);

        foldButton = new JButton("Convert a Folder");
        foldButton.addActionListener(this);

        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(fileButton);
        buttonPanel.add(foldButton);

        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
    }

    public void actionPerformed(ActionEvent e) 
    {

        //Handle open button action.
        if (e.getSource() == fileButton) 
        {
        	fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
        	int input = fc.showOpenDialog(WordMaker.this);

            if (input == JFileChooser.APPROVE_OPTION) 
            {
                File spreadsheet = fc.getSelectedFile();
                try 
                {
					if(FilenameUtils.getExtension(spreadsheet.getPath()).equals("xlsx"))
					{
						//This is where a real application would open the file.
						log.append("Accessing file: " + spreadsheet.getName() + "." + newline);
						
						fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						input = fc.showOpenDialog(WordMaker.this);
		    			if(input == JFileChooser.APPROVE_OPTION)
		    			{
		    				File dropoff = fc.getSelectedFile();
		    				FileOutputStream out = new FileOutputStream(new File(dropoff.getPath() + scanWrite(spreadsheet)) +".docx");
	    			        doc.write(out);
	    			        out.close();
		    			}
		    			else
		    				log.append("There is something wrong with this spreadsheet." + newline);
					}
					else
						log.append("File is not an Excel spreadsheet." + newline);
				} 
                catch (Exception e_) 
                {
                	log.append("Oof -" + e_);
				}
            } 
            log.setCaretPosition(log.getDocument().getLength());

        //Handle save button action.
        } 
        else if (e.getSource() == foldButton) 
        {
        	fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        	int input = fc.showOpenDialog(WordMaker.this);

            if (input == JFileChooser.APPROVE_OPTION) 
            {
                File folder = fc.getSelectedFile();
                File[] folderFiles = folder.listFiles();	
    			boolean[] isSheet = new boolean[folderFiles.length];
    			int sheetCount = 0;
    				
    			for(int i = 0; i < folderFiles.length; i++)
    				if(FilenameUtils.getExtension(folderFiles[i].getPath()).equals("xlsx"))
    				{
    					isSheet[i] = true;
    					sheetCount++;
    				}
    			
    			input = fc.showOpenDialog(WordMaker.this);
    			if(input == JFileChooser.APPROVE_OPTION)
    			{
	    			File dropoff = fc.getSelectedFile();	
	    			
	    			if(sheetCount > 0)
	    			{
		    			for(int i = 0; i < folderFiles.length; i++)
		    		        if(isSheet[i])
		    		        {
			    		       try
			    		       {
		    		        		FileOutputStream out = new FileOutputStream(new File(dropoff.getPath() + scanWrite(folderFiles[i])) +".docx");
			    			        doc.write(out);
			    			        out.close();
			    		       }
			    		       catch(Exception e_)
			    		       {
			    		    	   log.append("Oof -" + e_);
			    		       }
		    		        } 	
	    			}
					else
					{
						log.append("The folder contains no Excel spreadsheets. Please retry." + newline);
					}
    			}
    			else
    				log.append("I don't know how, but you broke it. (You shouldn't be seeing this.)" + newline);
            } 
            log.setCaretPosition(log.getDocument().getLength());
        }
    }

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    private static void createAndShowGUI() 
    {
        //Create and set up the window.
        JFrame frame = new JFrame(version);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new WordMaker());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) 
    {
        //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        SwingUtilities.invokeLater(new Runnable() 
        {
            public void run() 
            {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE); 
                createAndShowGUI();
            }
        });
    }
    
    private static String scanWrite(File spreadsheet) throws Exception //Scans an Excel spreadsheet and converts it into a Word document
	{
		//The next three lines of code first make the file readable, then treat it as an Excel workbook, and then finally target the first sheet.
		//Each Excel spreadsheet should only have content in one sheet, after all.
		FileInputStream ExcelFileToRead = new FileInputStream(spreadsheet);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        String collection = "";
        int loopCount = 1;
        int crashPortion = 0;
        XSSFRow currRow = wb.getSheetAt(0).getRow(0);
        
        try
        {
	        /*
	        	In the Apache POI (which we are using to work with Excel spreadsheets), a document, which is a workbook, is comprised of sheets.
	         	Sheets, in turn, are comprised of rows, which themselves are comprised of cells.
	         	In order to select a particular cell, we must choose a row, then a cell number.
	         	This is similar to a 2D Cartesian coordinate system, where each point can be denoted with two values, (x, y).
	         	The initialization of the two variables [row] and [cell] effectively act as these two values throughout <scanWrite>.
	        */
        	
        	// SECTION ZERO - INITIALIZATION OF SPREADSHEET AND DEFINITION OF COLUMNS
        	XSSFSheet sheet = wb.getSheetAt(0);
        	
	        XSSFRow row; 
	        XSSFCell cell;
			
	        //The variables below will denote the columns for the categories in the spreadsheet.
	        //In the case where the category does not exist, the value of its respective variable will be -1.
	        int collectionName = -1, collectionId = -1, accessionDate = -1, cont1 = -1, cont1Start = -1, cont1End = -1, 
	        	cont2 = -1, cont2Start = -1, cont2End = -1, groupSeries = -1, series = -1, subseries = -1, subsubseries = -1, heading = -1, 
	        	description = -1, medium = -1, form = -1, dateExpression = -1, namedEntities = -1, beginDate = -1, endDate = -1;
	        
	        //The first row of the sheet should contain all categories used in the sheet.
	        //With the next two lines of code, the first row is obtained, and a Cell iterator is created.
	        //The Cell iterator progress cell by cell, which is effectively moving column by column.
	        XSSFRow firstRow = sheet.getRow(0);
	        Iterator<Cell> firstCells = firstRow.cellIterator();
	        
	        //This portion of code is supposed to assign the categories to their respective columns.
	        //Note that if a cell has content of "Sub subseries", it will not update the [subsubseries] variable.
	        //It might be a good choice later on to disregard case.
	        while (firstCells.hasNext())
	        {
	            cell = (XSSFCell) firstCells.next();   
	            if (cell.getStringCellValue().equals("Collection Name"))
	                collectionName = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Collection Number"))
	                collectionId = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Accession Date"))
		            accessionDate = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Cont 1"))
	                cont1 = cell.getColumnIndex();  	
	            else if(cell.getStringCellValue().equals("Cont 1 Start"))
	                cont1Start = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Cont 1 End"))
	                cont1End = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Cont 2"))
	                cont2 = cell.getColumnIndex();           	
	            else if(cell.getStringCellValue().equals("Cont 2 Start"))
	                cont2Start = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Cont 2 End"))
	                cont2End = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equalsIgnoreCase("Grouping Series"))
	            	groupSeries = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Series"))
	                series = cell.getColumnIndex();             	
	            else if(cell.getStringCellValue().equals("Subseries"))
	                subseries = cell.getColumnIndex();            	
	            //Subsubseries must be written in the manner below.
	            else if(cell.getStringCellValue().equals("Sub-subseries"))
	                subsubseries = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Heading"))
	                heading = cell.getColumnIndex();           	
	            else if(cell.getStringCellValue().equals("Description"))
	                description = cell.getColumnIndex();           	
	            else if(cell.getStringCellValue().equals("Medium"))
	                medium = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Form"))
	                form = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Date Expression"))
	                dateExpression = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("Named Entities"))
	                namedEntities = cell.getColumnIndex();
	            else if(cell.getStringCellValue().equals("Begin Date"))
	                beginDate = cell.getColumnIndex();            	
	            else if(cell.getStringCellValue().equals("End Date"))
	                endDate = cell.getColumnIndex();
	        }
	        
	        /*
	        	The below four values are created for proper numbering purposes.
	        	[boxTracker] counts the boxes within each collection.
	        	[seriesTracker] counts the series within each box, numbered Roman numerically.
	        	[subseriesTracker] counts the subseries within each series, numbered alphabetically.
	        	[headingTracker] counts the items within each subseries (or series if subseries doesn't exist), numbered numerically.
	        */
	        
	        // SECTION ONE - ASSEMBLY OF HEADER
	        crashPortion = 1;
	        
	        int boxTracker;
	        String groupTracker = null;		//might not always exist
	        boolean sameGroup = true; 				//only used if groups exist
	        String seriesTracker;
	        String subseriesTracker;	//also might not always exist
	        String subsubseriesTracker = ""; //ditto
	        String headingTracker; 		//see above
	        
	        //The name of the collection is obtained by accessing the first row of data.
	        //Collection names should be the same throughout the spreadsheet.
	        //[collection] is printed out to show which spreadsheet the program is operating on, for debugging purposes.
	        collection = sheet.getRow(1).getCell(collectionName).getStringCellValue();
	        log.append("Now scanning: " + collection + newline);
	        
	        //[doc] and [content] are initialized. 
	        //[content] will be updated throughout the rest of the code until finally being put onto [doc] in the main method.
	        //[df] is implemented so that the String values of cells will be properly returned.
	        doc = new XWPFDocument();
	        content = new ArrayList<XWPFParagraph>();
	        DataFormatter df = new DataFormatter();
	
			//The heading for the document is created.
	        //The collection name, collection ID, and accession dates are properly printed.
	        makeNewRun("C", 0, 0);
			activeRun.setText(collection);
			activeRun.addBreak();
			activeRun.setText("" + df.formatCellValue(sheet.getRow(1).getCell(collectionId)));
			activeRun.addBreak();
			Iterator<Row> accessionRows = sheet.rowIterator();
			row = (XSSFRow) accessionRows.next();
			row = (XSSFRow) accessionRows.next(); //hopping to the second row, AKA the first content row
			
			//The following loop searches for all unique accession dates throughout the spreadsheet.
			String currentAccessionDate = df.formatCellValue(row.getCell(accessionDate));
			String accessionExpression = currentAccessionDate;
			while(!df.formatCellValue(row.getCell(accessionDate)).isEmpty())
			{
				if(!(currentAccessionDate.equals(df.formatCellValue(row.getCell(accessionDate))))) 
				{
					accessionExpression =  accessionExpression + ", " + df.formatCellValue(row.getCell(accessionDate));
					currentAccessionDate = df.formatCellValue(row.getCell(accessionDate));
				}
				if(accessionRows.hasNext())
					row = (XSSFRow) accessionRows.next();
				else
					break;
			}
			activeRun.setText(accessionExpression);	
			activeRun.addBreak();
			activeRun.setText("Preliminary Listing");	
			
			//The following counters are created for the proper labeling of each row.
			//As subseries and subsubseries are enumerated alphabetically, their types are set to char.
			
			// SECTION TWO - PRE-LOOP SETUP
			crashPortion = 2;
			
			int romanNum, itemNum;
			char subLetter, subsubLetter;
			
			boolean sameBox = true;
			
			//An iterator is created, and is set to start on the second row.
			Iterator<Row> rows = sheet.rowIterator();
			row = (XSSFRow) rows.next();
	    	row = (XSSFRow) rows.next();
	        
			log.append("Looping..." + newline);
			
			/*
				In general, the looping process is very recursive.
				The series of loops slowly work their way from descriptive to broad identifiers.
				Ordered from small-scale to large-scale, the ordering is: subsubseries, item, subseries, series, container.
				Subsubseries and subseries do not always exist.
			*/
		
	        //The termination condition for the while loop below is the absence of another row.
			//This should stop once there are no rows with content left.
			while(rows.hasNext() && !row.getCell(collectionName).getStringCellValue().isEmpty())
	        {	
	        	// SECTION THREE - CONTAINER 1 LAYER
				crashPortion = 3;
				
				currRow = row;
				
				if(groupSeries != -1)
				{
					groupTracker = row.getCell(groupSeries).getStringCellValue();
					
					makeNewRun("L", 0, 0);
					activeRun.setBold(true);
					activeRun.setText(groupTracker);
				}
				
				//[boxTracker] is updated to the proper box count.
				boxTracker = (int) row.getCell(cont1Start).getNumericCellValue();
	        	
				//The box and its number is written in the document.
				//Items can run across multiple boxes, so that must be checked.
				//Following the format, it will not be indented.
				makeNewRun("L", 0, 0);
	        	if(df.formatCellValue(row.getCell(cont1Start)).equals(df.formatCellValue(row.getCell(cont1End))))
					activeRun.setText(row.getCell(cont1) + " " + df.formatCellValue(row.getCell(cont1Start)));
				else 
					activeRun.setText(row.getCell(cont1) + " " + df.formatCellValue(row.getCell(cont1Start)) + "-" 
										+ df.formatCellValue(row.getCell(cont1End)));
				
	        	//The series count is reset to 1 for each box.
	        	romanNum = 1;
	        	
	        	//The termination condition for the while loop below is a change in box number.
	        	//This should stop once the selected item starts in a box value that is not [boxTracker].
	            while(row.getCell(cont1Start).getNumericCellValue() == boxTracker) 
	            {
	            	// SECTION FOUR - SERIES LAYER
	            	crashPortion = 4;
	            	
	            	//[seriesTracker] is updated to the current series.
	            	seriesTracker = row.getCell(series).getStringCellValue();
	            	
	            	//The series and its number (in Roman numerals) is written in the document.
	            	//Following the format, it will be indented once.
	            	makeNewRun("L", 1, 0);
	        		activeRun.setText(toRomanNum(romanNum++) + ". " + row.getCell(series).getStringCellValue());
	
	            	//The subseries letter is reset to 'A' for each series.
	        		subLetter = 65;
	            	
	        		//The termination condition for the while loop below is a change in series.
	        		//This should stop once the selected item is in a different series than [seriesTracker].
	        		while(row.getCell(series).getStringCellValue().equals(seriesTracker)) 
	            	{
	            		// SECTION FIVE - SUBSERIES LAYER
	        			
	        			//[subseriesTracker] is updated to the current subseries.
	        			subseriesTracker = row.getCell(subseries).getStringCellValue();
	        			
	        			if(subsubseries != -1)
	        			{
	        				subsubseriesTracker = row.getCell(subsubseries).getStringCellValue();	        					
	        			}
	        				   			
	            		//As it is possible for there not to exist a subseries, such is checked.
	        			//In this case, the items will just be printed.
	        			//Otherwise, the subseries and its letter is written in the document.
	        			//Following the format, it will be indented twice.
	        			if(!row.getCell(subseries).getStringCellValue().isEmpty())
	        			{
	        				makeNewRun("L", 2, 0);
	        				if(subsubseries != -1 && !row.getCell(subsubseries).getStringCellValue().isEmpty())
	        					activeRun.setText(subLetter++ + ". " + row.getCell(subseries).getStringCellValue() + " - " + row.getCell(subsubseries).getStringCellValue());
	        				else
	        					activeRun.setText(subLetter++ + ". " + row.getCell(subseries).getStringCellValue());
	        			}
	        			else
	        			{
	        				if(subsubseries != -1 && !row.getCell(subsubseries).getStringCellValue().isEmpty())
	        				{
	        					makeNewRun("L", 2, 0);
		        				activeRun.setText(subLetter++ + ". (No subseries) - " + row.getCell(subsubseries).getStringCellValue());
	        				}	        				
	        			}
	        				
	        					
	
	            		//The item count is reset to 1 for each item.
	        			itemNum = 1;
	        			
	        			//The termination condition for the while loop below is a change in subseries.
	        			//This should stop once the selected item is in a different subseries than [subseriesTracker].
	            		while(row.getCell(subseries).getStringCellValue().equals(subseriesTracker)) 
	            		{
	            			if(subsubseries != -1 && !row.getCell(subsubseries).getStringCellValue().equals(subsubseriesTracker))
	            				break;
	            			// SECTION SIX - ITEM LAYER: HEADER AND MOST DETAILS ASSEMBLY
	            			crashPortion = 6;
	            			
	            			//[loopCount] represents the total amount of items added in each document.
	            			//In the case of an exception being thrown, [loopCount] is reported.
	            			//This is so it is possible to go to the associated row in the spreadsheet in which the code failed, helping debug it.
	            			loopCount++;
	            			
	            			//Following the format, each item will be indented thrice.
	            			//Due to the format being like a list, the numbering of each item is considered.
	            			makeNewRun("L", 3, itemNum);

	            			//The following code parses the current row for the header, description, medium, form, and date expression.
	            			//In the case which a certain parameter is absent, extra punctuation will not be added.
	            			String headerAndDetails = "";
	            			String[] separatedDesc = null;
	            			if(!df.formatCellValue(row.getCell(heading)).isEmpty())
	            			{
	            				//Some headers have surrounding quotation marks.
	            				//We make maneuvers for the insertion of the comma
	            				String headerString = df.formatCellValue(row.getCell(heading));
	            				if(headerString.charAt(0) == '"') 
	            				{
	            					headerString = headerString.substring(0, headerString.length() - 1);
		            				
		            				if(!row.getCell(description).getStringCellValue().isEmpty())
		            					headerAndDetails = headerString + ",\" ";
		            				else
		            					headerAndDetails = headerString + "\" ";
	            				}
	            				else if(row.getCell(description) != null && !row.getCell(description).getStringCellValue().isEmpty())
		            					headerAndDetails = headerString + ", ";
	            				else
	            					headerAndDetails = headerString;
	            				
	            			}
	            			if(row.getCell(description) != null && !row.getCell(description).getStringCellValue().isEmpty())
	            			{
	            				separatedDesc = row.getCell(description).getStringCellValue().split(":");
	            				headerAndDetails = headerAndDetails + separatedDesc[0];
	            			}
	            			if(!row.getCell(medium).getStringCellValue().isEmpty())
	            				headerAndDetails = headerAndDetails + ", " + row.getCell(medium).getStringCellValue();
	            			if(!df.formatCellValue(row.getCell(form)).isEmpty())
	            				headerAndDetails = headerAndDetails + ", " + df.formatCellValue(row.getCell(form));
	            			
	            			// SECTION SEVEN - ITEM LAYER: DATE EXPRESSION AND CONTAINER 2 ASSEMBLY
	            			crashPortion = 7;
	            			
	            			if(!row.getCell(beginDate).toString().isEmpty())
	            			{	
	            				String begin = row.getCell(beginDate).toString();
	            				String end = row.getCell(endDate).toString();
	            				
	            				//The dates have to be formatted in a certain manner, so the <formatDate> function is supposed to do such.
	            				//This feature is still in testing, as we are uncertain of its consistency.
	            				headerAndDetails = headerAndDetails + "; " + formatDate(begin, end) + ".";
	            			}
	            			else
	            				headerAndDetails = headerAndDetails + "; n.d.";
	            			
	            			//The item is written into the document.
	            			activeRun.setText(itemNum++ + ". " + headerAndDetails);
	            			
	            			if(separatedDesc != null && separatedDesc.length == 2)
	            			{
	            				//The [subsubLetter] is reset to 'a' for each item.
	            				subsubLetter = 97;
	            				
	            				String[] individualEntities = separatedDesc[1].split(";");
	            				for(int i = 0; i < individualEntities.length; i++)
	                			{
	            					//Each named entity and its letter are written in the document.
	            	            	//Following the format, it will be indented four times.
	            					makeNewRun("L", 4, 0);
	                				activeRun.setText(subsubLetter++ + ". " + individualEntities[i] + ".");
	                			}
	            			}
	            			
	            			//The folder in which the item is contained is written into the document.
	            			//Items can run across multiple folders, so that condition much be checked.
	            			makeNewRun("R", 0, 0);
	            			char contentTwo = '?';
	            			if(!df.formatCellValue(row.getCell(cont2)).isEmpty())
	            				contentTwo = row.getCell(cont2).getStringCellValue().charAt(0);
	            			
	            			if(df.formatCellValue(row.getCell(cont2Start)).equals(df.formatCellValue(row.getCell(cont2End))))
	            				activeRun.setText("[" + contentTwo + ". " + df.formatCellValue(row.getCell(cont2Start)) + "]");
	            			else 
	            				activeRun.setText("[" + contentTwo + ". " + df.formatCellValue(row.getCell(cont2Start)) + "-" 
	            									+ df.formatCellValue(row.getCell(cont2End)) + "]");
	            			
	            			//[headingTracker] is updated for the current item.
	            			headingTracker = row.getCell(heading).toString();

	            			if(!rows.hasNext())
	            				break;
	            			
	            			row = (XSSFRow) rows.next();
	            			
	            			sameBox = row.getCell(cont1Start).getNumericCellValue() == (boxTracker);
	            			
	            			if(groupSeries != -1)
	            			{
	            				sameGroup = row.getCell(groupSeries).getStringCellValue().equals(groupTracker);
	            			}
	            			
	            			if(!sameGroup || !sameBox)
	            				break;
	            		}	//Exits subseries loop
	            		//Once again, more conditions to exit the appropriate loops.
	            		if(!rows.hasNext() || row.getCell(collectionName).getStringCellValue().isEmpty() || !sameGroup || !sameBox)
        					break;
	            	}	//Exits series loop
	        		if(!rows.hasNext() || row.getCell(collectionName).getStringCellValue().isEmpty() || !sameGroup || !sameBox)
    					break;
	            }	//Exits box loop
	        }	//Exits spreadsheet loop.
	        wb.close();
	        log.append("Program run without any errors." + newline);
        }
        catch(Exception e)
        {
        	//In case of a crash, the exception and row where the program failed will be printed.
        	//The document will not be created if this occurs.
        	log.append("CRASH! ERROR EXCEPTION: " + e + newline);
        	log.append("Program crashed on row number: " + loopCount + newline);
        	log.append("> Further Information for Debugging... <" + newline);
        	StackTraceElement[] stktr = e.getStackTrace();
			for(int i = 0; i < stktr.length; i++)
				log.append(stktr[i].toString() + newline);
        	switch(crashPortion)
        	{
        		case 0:
        		{
        			log.append("Crashed in the initialization - check for whether the spreadsheet actually had information." + newline);
        			break;
        		}
        		case 1:
        		{
        			log.append("Crashed in the header assembly - perhaps a field row was improperly named" + newline);
        			break;
        		}
        		case 2:
        		{
        			log.append("Crashed in the pre-loop assembly - ensure that the spreadsheet has more than one row" + newline);
        			break;
        		}
        		case 3:
        		{
        			log.append("Crashed in the Container 1 column" + newline);
        			break;
        		}
        		case 4:
        		{
        			log.append("Crashed in the Series column" + newline);
        			break;
        		}
        		case 5:
        		{
        			log.append("Crashed in the Subseries column" + newline);
        			break;
        		}
        		case 6:
        		{
        			log.append("Crashed in the Header, Description, Medium, or Named Entities column" + newline);
        			break;
        		}
        		case 7:
        		{
        			log.append("Crashed in the Date Expression or Container 2 column - check whether the dates are in the proper format" + newline);
        			break;
        		}
        		case 8:
        		{
        			log.append("Crashed in the Subsubseries column" + newline);
        			break;
        		}
        		default:
        		{
        			log.append("Glitch! You shouldn't be seeing this!" + newline);
        		}
        	}
        }
        
        //The title will be returned so the document can be named such.
        return collection;
	}
	
	private static void makeNewRun(String pAlign, int indentFactor, int bulletNumber) //Creates and formats a new paragraph
	{
		/*
			<makeNewRun> creates a new paragraph within [content].
			First, a paragraph is created in [doc], and then added to [content].
			[activePara] is then properly set to the new paragraph.
			The paragraph is then formatted as prompted:
				[indentFactor] denotes the quantity of 0.5" indents.
				[bulletNumber] is used for the neatness of items.
			Finally, after being spaced, a run is created and [activeRun] is set to it.
			The font throughout the entire document will be Times New Roman.
		*/
		content.add(doc.createParagraph());
    	activePara = content.get(content.size() - 1);
    	
    	//Aligns the current [activePara] as prompted.
    	//"L" means left, "C" means right, and "R" means right.
    	//An improper input will be defaulted to left alignment, and an error message will be printed.
    	if(pAlign.toUpperCase().equals("L"))
    	{
    		activePara.setAlignment(ParagraphAlignment.LEFT);
    		
    		//As the Apache POI indents by pixels instead of inches, a conversion rate is necessary.
    		//In the case of adding items, an extra hanging indent is added to follow the format, as items typically occupy over one line.
    		//In standard list numbering, the item words should never be directly below its numbering.
    		activePara.setIndentationLeft((indentFactor * 360) + ((digits(bulletNumber) + 2) * 90)); //720 unit = 0.5 inch
        	activePara.setIndentationHanging((digits(bulletNumber) + 2) * 90);
    	}
    	else if(pAlign.toUpperCase().equals("R"))
    		activePara.setAlignment(ParagraphAlignment.RIGHT);
    	else if(pAlign.toUpperCase().equals("C"))
    		activePara.setAlignment(ParagraphAlignment.CENTER);
    	else
    		log.append("Improper Format! Defaulting to LEFT Alignment.");
    	activePara.setSpacingAfter(80);
    	activePara.createRun();
    	activeRun = activePara.getRuns().get(0);
    	activeRun.setFontFamily("Times New Roman");
	}
	
	private static String toRomanNum(int i) //supports all integers from [1, 50)
	{
		//This follows the principles of Roman numerals and constructs a string that represents the inputed integer.
		String ret = "";
		if(i / 10 == 4)
			ret = ret + "XL";
		else
			for(int j = 0; j < (i / 10); j++)
				ret = ret + "X";
		if(i % 5 == 4)
		{
			if(i % 10 == 9)
				ret = ret + "IX";
			else
				ret = ret + "IV";
		}
		else
		{
			if((i / 5) % 2 == 1)
				ret = ret + "V";
			for(int j = 0; j < (i % 5); j++)
				ret = ret + "I";
		}
		return ret;
	}
	
	private static int digits(int i) //supports all integers from [0, 1000)
	{
		//This returns the number of digits that a number has.
		int ret = 0;
		if(i / 1 > 0)
			ret++;
		if(i / 10 > 0)
			ret++;
		if(i / 100 > 0)
			ret++;
		return ret;
	}
	
	private static String formatDate(String begin, String end) 
	{
		String[] seasonOptions = {"Spring", "Summer", "Fall", "Winter", "Spring/Summer", "Fall/Winter"};
		String[] monthOptions = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
									"November", "December"};
		String begYear = "", begMonth, begDay, endYear = "", endMonth, endDay;
		
		SimpleDateFormat slashForm = new SimpleDateFormat("yyyy/mm/dd");
		SimpleDateFormat dashForm = new SimpleDateFormat("yyyy-mm-dd");
		
		if(end == null || end.length() == 0)
			return begin;
		if(begin.length() > 7) 
		{
			//check if the years are the same and check if the date is the first of the month
			begYear = begin.substring(0, 4);
			begMonth = begin.substring(5, 7);
			begDay = begin.substring(begin.length() - 2);
			endYear = end.substring(0, 4);
			endMonth = end.substring(5, 7);
			endDay = end.substring(begin.length() - 2);
		}
		else if(begin.length() > 4)
		{
			//when a year isn't given
			begMonth = begin.substring(2, 4);
			begDay = begin.substring(begin.length() - 2);
			endMonth = end.substring(2, 4);
			endDay = end.substring(begin.length() - 2);
		}
		else
		{
			if(begin.equals(end))
				return begin;
			else
				return begin + "-" + end;
		}
		
		//Do all the checks that require the beginning and end years to be the same
		if(begYear.equals(endYear) && begDay.equals("01")) {
			//Checking for seasons
			if(begMonth.equals("03")  && endMonth.equals("05") && endDay.equals("31")) {
				return seasonOptions[0] + " " + begYear;
			}else if(begMonth.equals("06")  && endMonth.equals("08") && endDay.equals("31")) {
				return seasonOptions[1] + " "  + begYear;
			}else if(begMonth.equals("09")  && endMonth.equals("11") && endDay.equals("30")) {
				return seasonOptions[2] + " "  + begYear;
			}else if(begMonth.equals("12")  && endMonth.equals("02") && (endDay.equals("29") || endDay.equals("28"))) {
				return seasonOptions[3] + " "  + begYear;
			}else if(begMonth.equals("03")  && endMonth.equals("08") && endDay.equals("31")) {
				return seasonOptions[4] + " "  + begYear;
			}else if(begMonth.equals("09")  && endMonth.equals("02") && (endDay.equals("29") || endDay.equals("28"))) {
				return seasonOptions[5] + " "  + begYear;
				
			//Check for the whole year
			}else if(begMonth.equals("01")  && endMonth.equals("12") && (endDay.equals("31"))) {
				return "ca. " + begYear;
			
			//Checking for months
			}else if(begMonth.equals(endMonth)) {
				if(begMonth.equals("01") && endDay.equals("31")) {
					return monthOptions[0] + " "  + begYear;
				}else if(begMonth.equals("02") && ((endDay.equals("28") || endDay.equals("29")))){
					return monthOptions[1] + " "  + begYear;
				}else if(begMonth.equals("03") && (endDay.equals("31"))){
					return monthOptions[2] + " "  + begYear;
				}else if(begMonth.equals("04") && (endDay.equals("30"))){
					return monthOptions[3] + " "  + begYear;
				}else if(begMonth.equals("05") && (endDay.equals("31"))){
					return monthOptions[4] + " "  + begYear;
				}else if(begMonth.equals("06") && (endDay.equals("30"))){
					return monthOptions[5] + " "  + begYear;
				}else if(begMonth.equals("07") && (endDay.equals("31"))){
					return monthOptions[6] + " "  + begYear;
				}else if(begMonth.equals("08") && (endDay.equals("31"))){
					return monthOptions[7] + " "  + begYear;
				}else if(begMonth.equals("09") && (endDay.equals("30"))){
					return monthOptions[8] + " "  + begYear;
				}else if(begMonth.equals("10") && (endDay.equals("31"))){
					return monthOptions[9] + " "  + begYear;
				}else if(begMonth.equals("11") && (endDay.equals("30"))){
					return monthOptions[10] + " "  + begYear;
				}else if(begMonth.equals("12") && (endDay.equals("31"))){
					return monthOptions[11] + " "  + begYear;
				}
				
			}
			
		
		
		//Do the checks that require the beginning and end years to be different
		}else if(begMonth.equals("01")&& begDay.equals("01") && endMonth.equals("12") && endDay.equals("31")) {
			if(Integer.parseInt(endYear) - Integer.parseInt(begYear) == 10 ) {
				return begYear.substring(0,2) + "00s";
			}else {
				return begYear + "-" + endYear;
			}
			
		//now if the dates don't fit any of the criteria, print them normally

		}
		String ret = "";
		
		//need to format differently if there isn't a year
		if(begin.length() < 8) {
			ret = "--" + begMonth + "--" + begDay;
			if(!end.isEmpty()) {
				ret = ret + "-" + endMonth + "--" + endDay;
			}
			return ret;
		}
		try 
		{	
			ret = slashForm.format(dashForm.parse(begin));
			if(!begin.equals(end) && !end.isEmpty())
				ret = ret + "-" + slashForm.format(dashForm.parse(end));
		}
		catch(ParseException parse)
		{
			if(begin.substring(begin.length() - 2).equals(".0"))
				begin = begin.substring(0, begin.length() - 2);
			if(!end.isEmpty() && end.substring(end.length() - 2).equals(".0"))
			{
				end = end.substring(0, end.length() - 2);
			}
			ret = begin;
			if(!begin.equals(end) && !end.isEmpty())
				ret = ret + "-" + end;
		}
		return ret;
	}
}