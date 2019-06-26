package DataDriven_Test;

import java.io.IOException;
import java.util.ArrayList;

public class SampleTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
	
		DOBExcelSheetTest obj = new DOBExcelSheetTest();
		
		ArrayList<String> data = obj.getData("April");
		
		System.out.println("the string at index 0 is ="+ data.get(0));

		System.out.println("the string at index 1 is ="+data.get(1));
	}

}
