import java.io.File;
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
	public static void main(String[] args)
	{
		//Scanner reader = new Scanner(System.in);
		//System.out.println(reader.nextLine());
		
		org.apache.log4j.PropertyConfigurator.configure("C:\\Users\\student\\Desktop\\Tess4J/log4j.properties.txt");		
		
		File image = new File("C:\\Users\\student\\Desktop\\Hatchet/Blast.pdf");
		Tesseract inst = new Tesseract();

		inst.setTessVariable("preserve_interword_spaces", "1"); //does something visible
		inst.setTessVariable("gapmap_use_ends", "1"); //does nothing visible
		inst.setTessVariable("tessedit_create_hocr", "1");
		
		//inst.setHocr(true);
		inst.setPageSegMode(12);
		
		try
		{
			String ret = inst.doOCR(image);
			System.out.println(ret);
		}
		catch(TesseractException e)
		{
			System.err.println(e.getMessage());
		}
		
		// END
	}
}
