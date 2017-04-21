package tw.com;
//Author Iris Chen
//extract android part of original file and save file as subname .xls

import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFRow;

public class MainPath {
	public static void main (String[] args)
	{
		ExcelParseToXml mExcelParseToXml = new ExcelParseToXml ();
	    
		String xlsPath ="C://Users//iris//Desktop//OMG_String.xls";
//		String xlsPath ="C://Users//iris//Desktop//String_test.xls";
	    mExcelParseToXml.execute(xlsPath);
	    
	}
}
