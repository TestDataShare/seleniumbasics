package readfromexcel;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class program1 {
	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\selenium\\Excel\\dataread.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int row = sheet.getLastRowNum();
		int cell = sheet.getRow(0).getLastCellNum();
		System.out.println(row);
		System.out.println(cell);
		for(int r=0;r<=row; r++)
		{
			XSSFRow rowno = sheet.getRow(r);
			for(int c=0; c<cell; c++)
			{
				String cellval = rowno.getCell(c).toString();
				System.out.print(cellval+"    ");
			}
			System.out.println();
		}
		workbook.close();
		file.close();
	}
}
