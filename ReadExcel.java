package writereadexcel;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

//Locating the workbook & sheet
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\PushpavathyNaveen\\eclipse-workspace\\ExcelReader\\Readfile.xlsx");
		XSSFSheet sheet = book.getSheetAt(0);
		
//Getting the row count 
		int rowcount = sheet.getLastRowNum();
//Getting column count
		int columncount = sheet.getRow(0).getLastCellNum();
		String[][] data = new String[rowcount][columncount];
		
//Loop to get value
		for ( int i=1; i<=rowcount;i++) {
			
			XSSFRow row = sheet.getRow(i);
		
			for(int j=0;j< columncount;j++) {
			
				XSSFCell cell = row.getCell(j);
				
				System.out.println(cell.getStringCellValue());
		
			}
		
	}
book.close();
	}
}
