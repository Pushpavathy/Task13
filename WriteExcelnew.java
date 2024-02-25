package writereadexcel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelnew {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		//Create the workbook

		XSSFWorkbook book = new XSSFWorkbook();
		//Creating sheet
		XSSFSheet sheet = book.createSheet();

		// Getting the values to be inputed in the excel

		Object[][] data ={
				{"Name","Age","Email"},
				{"John Doe",30,"john@test.com"},
				{"Jane Doe",28,"john@test.com"},
				{"Bob Smith",35,"jacky@example.com"},
				{"Swapnil",37,"swapnil@example.com"}
		};
		//Create row

		int rowcount= 0;

		for (Object[] row : data)
		{
			XSSFRow row1 = sheet.createRow(rowcount++);

			//Create cell
			int colcount =0;

			for(Object col : row)
			{
				XSSFCell cell =row1.createCell(colcount++);

				//Type casting data	
				{if (col instanceof String)
				{
					cell.setCellValue((String)col);
				}
				else if (col instanceof Integer)
				{
					cell.setCellValue((Integer) col);
				}


				}

			}
		}

		//Creating and writing excel
		FileOutputStream output = new FileOutputStream("Newfile.xlsx");

		book.write(output);
	}

}
