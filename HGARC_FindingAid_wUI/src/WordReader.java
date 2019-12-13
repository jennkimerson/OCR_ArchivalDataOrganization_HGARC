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

public class WordReader 
{
	public static void main(String[] args) throws IOException
	{
		org.apache.log4j.PropertyConfigurator.configure("C:\\Users\\student\\Desktop\\Tess4J/log4j.properties.txt");
		//File full = new File("Z:\\PDF Sheets/Paul Brodeur Inventory Complete.pdf");
		File full = new File("Z:\\hgarc_fadi_faSamples/Arrighi, Mel 1980-1987.pdf");
		PDDocument pdfull = PDDocument.load(full);
		
		Splitter split = new Splitter();
		ArrayList<PDDocument> pages = (ArrayList<PDDocument>) split.split(pdfull);
		System.out.println(pages.size());		
		
		//File image = new File("C:\\Users\\student\\Desktop\\Hatchet/sampleArc1_edit.png");
		File current;
		ITesseract inst = new Tesseract();
		//ArrayList<String> configs = new ArrayList<String>();
		//configs.add("hocr");
		//inst.setConfigs(configs);
		inst.setTessVariable("preserve_interword_spaces", "1");
		inst.setLanguage("eng");
		//inst.setTessVariable("tessedit_create_hocr", "1");
		inst.setPageSegMode(5); //PSM_SINGLE_BLOCK_VERT_TEXT
		
		try
		{
			String[] ret = new String[pages.size()];
			for(int i = 0; i < 3/*pages.size()*/; i++)
			{
				pages.get(i).save("C:\\Users\\student\\Desktop\\Hatchet\\PBIC/page" + i + ".pdf");
				current = new File("C:\\Users\\student\\Desktop\\Hatchet\\PBIC/page" + i + ".pdf");
				ret[i] = inst.doOCR(current);
			}
			System.out.println(ret[1]);
		}
		catch(TesseractException e)
		{
			System.err.println(e.getMessage());
		}
	}
}
