// reads in the text vertically

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.Iterator;
import java.awt.List;
import java.io.IOException;

import net.sourceforge.tess4j.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;

public class Lab 
{
	private static XSSFWorkbook workb;
	private static int rowCount;
	private static int cellCount;
	private static XSSFRow currRow;
	private static XSSFCell currCell;
	
	public static void main(String[] args) throws IOException
	{
		//Scanner reader = new Scanner(System.in);
		//System.out.println(reader.nextLine());
		
		org.apache.log4j.PropertyConfigurator.configure("sample.txt");		
		
		File image = new File("sample.pdf");
		Tesseract inst = new Tesseract();

		inst.setTessVariable("preserve_interword_spaces", "1"); //does something visible
		inst.setTessVariable("gapmap_use_ends", "1"); //does nothing visible
		inst.setTessVariable("tessedit_create_hocr", "1");
		
		//inst.setHocr(true);
		inst.setPageSegMode(5); //PSM_SINGLE_COLUMN_VERT_TEXT
		
		String words = "";
		
		try
		{
			String ret = inst.doOCR(image);
			System.out.println(ret);
			words = ret;
		}
		catch(TesseractException e)
		{
			System.err.println(e.getMessage());
		}
		
		String[] pirate = words.split("\n");
		System.out.println(pirate.length);
		
		workb = new XSSFWorkbook();
		XSSFSheet sheetOne = workb.createSheet();
		rowCount = 0;
		
		for(int i = 0; i < pirate.length; i++)
		{
			currRow = sheetOne.createRow(rowCount++);
			currCell = currRow.createCell(cellCount);
			currCell.setCellValue(pirate[i]);
		}
		
		FileOutputStream out = new FileOutputStream(new File("result.xlsx"));
		workb.write(out);

		// END
	}
}
