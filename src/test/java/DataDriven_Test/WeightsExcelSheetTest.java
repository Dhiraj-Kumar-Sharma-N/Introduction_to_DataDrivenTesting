package DataDriven_Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WeightsExcelSheetTest {
	
	public ArrayList<String> getDatawithWeight(String Weight) throws IOException {
		
			ArrayList<String> a = new ArrayList<String>();
			
		FileInputStream fileDir = new FileInputStream("C://UDEMY_Selenium/Exceldata.xlsx");
		
		XSSFWorkbook WB = new XSSFWorkbook(fileDir);
		
		int sheetnum = WB.getNumberOfSheets();
		
		System.out.println(sheetnum);
		
		int k = 0;
		
		int columnid= 0;
		
		for (int i = 0; i < sheetnum; i++) {
			
			String sheetname = WB.getSheetName(i);
			
			System.out.println(sheetname);
			
			if (sheetname.equalsIgnoreCase("WeightDetails")) {
				
				XSSFSheet sheet = WB.getSheetAt(i);
				
				Iterator<Row> Rowiterator = sheet.iterator();
				
				Row row = Rowiterator.next();
				
				
				Iterator<Cell> cellvalue = row.cellIterator();
				
			while (cellvalue.hasNext()) {
				
				Cell cell = (Cell) cellvalue.next();
				
				System.out.println(cell.getStringCellValue());
				
				
				if (cell.getStringCellValue().equalsIgnoreCase("Weight")) {
					
					columnid = k;
					
					System.out.println("column id is = "+columnid);
											
						}
				k++;
						
					}
			while (Rowiterator.hasNext()) {
				
				Row rowite = (Row) Rowiterator.next();
				
				
				
				if (rowite.getCell(columnid).getStringCellValue().equalsIgnoreCase(Weight)) {
					
					Iterator<Cell> cellite = rowite.cellIterator();
					
					while (cellite.hasNext()) {
						
						Cell cell = (Cell) cellite.next();
						
						if (cell.getCellTypeEnum()==CellType.STRING) {
							
							
							a.add(cell.getStringCellValue());
							
						}
						
						else {
							
							a.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
						}
						
					}
					
				}
					
				}
				
			}
			}
		return a;	
	
	}
	
public static void main(String[] args) throws IOException {
	
	WeightsExcelSheetTest obj = new WeightsExcelSheetTest();
	
	ArrayList<String> data = obj.getDatawithWeight("hundred");
	
	System.out.println("the string at index 0 is ="+ data.get(0));

	System.out.println("the string at index 1 is ="+data.get(1));
	
	System.out.println("the string at index 2 is ="+data.get(2));
}

}

	


