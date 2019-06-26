package DataDriven_Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DOBExcelSheetTest {
	
	

	public  ArrayList<String> getData(String TestCaseName) throws IOException{
		// TODO Auto-generated method stub

		ArrayList<String> arr = new ArrayList<String>();
		
		FileInputStream Fis = new FileInputStream("C://UDEMY_Selenium/Exceldata.xlsx"); // specifying the file path using File input stream class 
		
		XSSFWorkbook workbook =new XSSFWorkbook(Fis); //Class used to retrieve the Workbook/ExcelSheet
		
		int NumofSheets = workbook.getNumberOfSheets(); //get number of sheets present in the ExcelSheet
		
		System.out.println(NumofSheets);
		
		for (int i = 0; i < NumofSheets; i++) {
			
			String SheetName =  workbook.getSheetName(i); //Get Sheet name for each of the sheet in Excel
			
			System.out.println(SheetName); //print naem of each sheet
			

if (SheetName.equalsIgnoreCase("DOBDetails")) { //Checks if the sheet name is equal to DOBDetails
	
	XSSFSheet Sheet= workbook.getSheetAt(i);  //get the sheet having that corresponding name
	
	
	// Identify TestCaseNumber column by scannig the entire 1st row (ROW HEADER)
	
	Iterator<Row> FirstRow = Sheet.iterator(); //sheet is a collection of Rows
	
	Row HeaderRowData = FirstRow.next();   //Rows are colllection  of Cells 
	
	Iterator<Cell>  RowCells = HeaderRowData.cellIterator();
	
	int k =0 ;
	
	int columnnumber =0 ;
	
	while (RowCells.hasNext()) {
		
		Cell cell = (Cell) RowCells.next();
		
		String cellvalue =  cell.getStringCellValue();
		
		System.out.println(cellvalue);
		
		if (cellvalue.equalsIgnoreCase("BirthMonth")) {
			
			columnnumber=k;
			
			System.out.println("header found at column number = "+columnnumber);
		}
			k++;
	}
	
	//Once the BirthMonth column is identified then scan the entire column for April row 	
	
	while (FirstRow.hasNext()) {
		
		Row rowScan = FirstRow.next();
				
		if(rowScan.getCell(columnnumber).getStringCellValue().equalsIgnoreCase(TestCaseName)) {
			
			System.out.println(rowScan.getRowNum());
			
	//Scan the entire column once April is found in the specified column
			
			
		Iterator<Cell> realCelldata =	rowScan.cellIterator();
		
		while (realCelldata.hasNext()) {
			Cell cell = (Cell) realCelldata.next();
			
			
			if (cell.getCellTypeEnum()==CellType.STRING) {
				
				arr.add(cell.getStringCellValue());
			}
			
			else {
				
				arr.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
			}
			
		
			
		}
			
		}
		
	}

	
	
}

//
			
			}
		return arr;
		
			
		}
	
	}
	


