package tw.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.ListIterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class ExcelParseToXml {

	private String mFileName = null;
	ArrayList<HSSFRow> rowdata = new ArrayList<HSSFRow>();
	private int totalStringRow = 0;
	private int totalStringArrayRow = 0;
	private int totalData = 0;

	public void execute(String xlsPath){
		int cellNum = ReadFromExcel (xlsPath);
		for(int i=3; i<cellNum; i++){
			WriteToXml(i);
		}
//		WriteToXml(3);
	}

	private int ReadFromExcel(String xlsPath) {
		 InputStream inputStream = null; 
		 HSSFRow row = null;
         int rowNumber = 0;
		 
         try
		    {
		        inputStream = new FileInputStream (xlsPath);
		    }
		    catch (FileNotFoundException e)
		    {
		        System.out.println ("File not found in the specified path.");
		        e.printStackTrace ();
		    }
		    POIFSFileSystem fileSystem = null;
		    try {
		    	
		        fileSystem = new POIFSFileSystem (inputStream);
		        HSSFWorkbook      workBook = new HSSFWorkbook (fileSystem);
		        HSSFSheet         sheetString    = workBook.getSheetAt (0); // read string 
		        HSSFSheet         sheetStringArray   = workBook.getSheetAt (1); // read string-array 
		        
		        Iterator<?> rowsString     = sheetString.rowIterator ();
		        Iterator<?> rowsStringArray     = sheetStringArray.rowIterator ();
		       
		        //extract file name
		        row = (HSSFRow) rowsString.next();
		        rowdata.add(row);
		        
		        rowsString.next();//remove the second row
		        
		        //process string
		        while (rowsString.hasNext ()) 
		        {
		            row = (HSSFRow) rowsString.next(); 

		            rowNumber = row.getRowNum ();
		            // display row number
		            System.out.println ("Row No.: " + rowNumber);

		            // remove useless cell from every row
		            row.removeCell(row.getCell(0));
		            row.removeCell(row.getCell(2));
		            
		            rowdata.add(row);
		            
		            if(row.getCell(3).getCellType()==Cell.CELL_TYPE_NUMERIC) {
	                	row.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
	                	
	                	}
		            System.out.println (row.getCell(1).getStringCellValue()+" "+row.getCell(3).getStringCellValue());
	            }	//end while
		        totalStringRow = rowNumber;
		        
		        rowsStringArray.next(); //remove first row
		        rowsStringArray.next(); //remove second row
		        while (rowsStringArray.hasNext ()) 
		        {
		        	row = (HSSFRow) rowsStringArray.next(); 
		        	
		        	rowNumber = row.getRowNum ();
		        	// display row number
		        	System.out.println ("Row No.: " + rowNumber);
		        	
		        	// remove useless cell from every row
		        	row.removeCell(row.getCell(0));
		        	row.removeCell(row.getCell(2));
		        	
		        	rowdata.add(row);
		        	if(row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC) {
	                	row.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
		        	}
		        	System.out.println (row.getCell(1).getStringCellValue()+" "+row.getCell(3).getStringCellValue());
		        } //end while
		        totalStringArrayRow = rowNumber;
		        totalData = totalStringRow+totalStringArrayRow;
		        System.out.println("totalStringRow=" + totalStringRow+" totalStringArrayRow="+totalStringArrayRow+" totalData="+totalData);
		    }
		    catch(IOException e)
		    {
		        System.out.println("IOException " + e.getMessage());
		    } 
		    System.out.println("getLastCellNum="+row.getLastCellNum());
		    return row.getLastCellNum();
		    
	}

	private void WriteToXml(int i) {
		try {
			// Initializing the XML document
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document document = builder.newDocument();
			Element rootElement = document.createElement("resources");
			
			document.appendChild(rootElement);
			
			ListIterator<HSSFRow> cells = rowdata.listIterator();
			
			HSSFRow cell = (HSSFRow)cells.next();
			mFileName = cell.getCell(i).toString();
		    createFolder();
		    System.out.println("i="+i+" cell.getCell(i).toString()="+cell.getCell(i).toString());
			
		    while(cells.hasNext()){
		    	int num = cells.nextIndex();
				cell = (HSSFRow)cells.next();
				
				if(cell.getCell(i).getCellType() == Cell.CELL_TYPE_NUMERIC) {
					cell.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                	System.out.println (cell.getCell(1).getStringCellValue()+" "+cell.getCell(i).getStringCellValue());}
				
				String mCellValue = cell.getCell(i).getStringCellValue();
				
				if(!cell.getCell(1).getStringCellValue().equals("")&&!mCellValue.equals("")){ //<string></string>
					
					Element stringElement = document.createElement("string");
					stringElement.setAttribute("name", cell.getCell(1).getStringCellValue());
					stringElement.appendChild(document.createTextNode(mCellValue));
					rootElement.appendChild(stringElement);
					
					System.out.println("String num="+num+" "+mCellValue);
					
				}else if(!cell.getCell(1).getStringCellValue().equals("")&&mCellValue.equals("")){  //<string-array></string-array>
					if(num < totalStringRow){ // to avoid no data in string part
						System.out.println("continue no string data num="+num+" getStringCellValue()="+mCellValue);
						continue;
					}
					if((num >= totalStringRow)&&(num<totalData)){ // to avoid no data in string-array part
						num = cells.nextIndex();
						HSSFRow cell_item = (HSSFRow)cells.next();
						

						if(cell_item.getCell(i).getCellType() == Cell.CELL_TYPE_NUMERIC) {
							cell_item.getCell(i).setCellType(Cell.CELL_TYPE_STRING);}
							
						if(cell_item.getCell(i).getStringCellValue().equals("")){
							System.out.println("continue no item data num="+num+" getStringCellValue()="+mCellValue);
							continue;
						}else{
							cells.previous();
						}
					}
					
					System.out.println("String-array num="+num+" "+mCellValue);
					
					Element stringArrayElement = document.createElement("string-array");
					stringArrayElement.setAttribute("name", cell.getCell(1).getStringCellValue());
					rootElement.appendChild(stringArrayElement);
					
					while(cells.hasNext()){
						num = cells.nextIndex();
						HSSFRow cell_item = (HSSFRow)cells.next();
						
						if(cell_item.getCell(i).getCellType() == Cell.CELL_TYPE_NUMERIC) {
							cell_item.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
		                	System.out.println (cell_item.getCell(1).getStringCellValue()+" "+cell_item.getCell(i).getStringCellValue());}
						
						String mCellItemValue = cell_item.getCell(i).getStringCellValue();
						
						System.out.println("String-array num="+num+" "+mCellItemValue);
						
						if(!mCellItemValue.equals("")&&cell_item.getCell(1).getStringCellValue().equals("")){  //<item></item>
							Element itemElement = document.createElement("item");
							itemElement.appendChild(document.createTextNode(mCellItemValue));
							stringArrayElement.appendChild(itemElement);
						}else if(mCellItemValue.equals("")&&(!cell_item.getCell(1).getStringCellValue().equals(""))){  // next string-array
							 cells.previous();
							 break;
						}else{ // data end
							break;
						}
						
					}
				}
			}
			
			TransformerFactory tFactory = TransformerFactory.newInstance();
			
			Transformer transformer = tFactory.newTransformer();
			// Add indentation to output
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "5");
			
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(new File(".//CopyMeToResourceFile//"+mFileName+"//strings.xml"));
			transformer.transform(source, result);
			
		} catch (ParserConfigurationException e) {
			System.out.println("ParserConfigurationException " + e.getMessage());
		} catch (TransformerConfigurationException e) {
			System.out.println("TransformerConfigurationException "+ e.getMessage());
		} catch (TransformerException e) {
			System.out.println("TransformerException " + e.getMessage());
		}
	}

	
	private void createFolder(){
		
		File newFolder = new File(".//CopyMeToResourceFile//"+mFileName);
		
			System.out.println("creating dir "+newFolder.toString());
			boolean result = false; 
			
			try{
				newFolder.mkdirs();
				result = true;
			}catch(Exception e){
				System.out.println(e);
			}
			
			if(result){
				System.out.println(newFolder.getName()+" is created");
			}
		
	}
}

