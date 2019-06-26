package DataDriven_Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoginExcelSheetTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream Fileinst = new FileInputStream("C://UDEMY_Selenium/Exceldata.xlsx");		
		
		XSSFWorkbook Workbook = new XSSFWorkbook(Fileinst);
		
		int NumberofSheets = Workbook.getNumberOfSheets();
		
		System.out.println(NumberofSheets);
		
		for (int i = 0; i < NumberofSheets; i++) {
			
			System.out.println(Workbook.getSheetName(i));
			
			if (Workbook.getSheetName(i).equalsIgnoreCase("LoginDetails")) {

				XSSFSheet sheet  =	Workbook.getSheetAt(i);
				
				Iterator<Row> headerRow = sheet.iterator();
				
				Row HeaderRowData = headerRow.next(); 
							
					
					Iterator<Cell>  cellvalue = HeaderRowData.cellIterator();
					
					int k =0;
					int columnNum = 0;
					
					while (cellvalue.hasNext()) {
						
						Cell cell = (Cell) cellvalue.next();
						
					
					String cellval =cell.getStringCellValue();
					
					if (cellval.equalsIgnoreCase("Username")) {
						
						columnNum =k;
					
					System.out.println(cell.getStringCellValue());
					
					System.out.println("header found at column number = "+columnNum);
									
				}
					
				k++;
			}
				
while (headerRow.hasNext()) {
	
	Row rowScan = (Row) headerRow.next();
	
	if (rowScan.getCell(columnNum).getStringCellValue().equalsIgnoreCase("Dhiraj")) {
		
		System.out.println(rowScan.getRowNum());
		
		
		Iterator<Cell> realCelldata =	rowScan.cellIterator();
		
		while (realCelldata.hasNext()) {
			Cell cell = (Cell) realCelldata.next();
			
			System.out.println(cell.getStringCellValue());
			
		}
		
	}
	
}

		
		
					
			break;
		}
	}

	}}

