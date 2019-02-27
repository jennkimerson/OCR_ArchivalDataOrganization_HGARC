import java.io.File;
import java.util.ArrayList;
import java.util.Scanner;

import net.sourceforge.tess4j.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class WordReader 
{
	public static void main(String[] args) 
	{
		org.apache.log4j.PropertyConfigurator.configure("C:\\Users\\student\\Desktop\\Tess4J/log4j.properties.txt");
		File image = new File("Z:\\PDF Sheets/Paul Brodeur Inventory Complete.pdf");
		//File image = new File("C:\\Users\\student\\Desktop\\Hatchet/sampleArc1_edit.png");
		ITesseract inst = new Tesseract();
		ArrayList<String> configs = new ArrayList<String>();
		//configs.add("hocr");
		//inst.setConfigs(configs);
		inst.setTessVariable("preserve_interword_spaces", "1");
		//inst.setTessVariable("tessedit_create_hocr", "1");
		
		try
		{
			String ret = inst.doOCR(image);
			System.out.println(ret);
		}
		catch(TesseractException e)
		{
			System.err.println(e.getMessage());
		}
	}
}
