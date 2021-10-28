package apachePoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Set {
	public static void main(String[] args) throws Exception {
// take control of the file
		File f = new File("C:\\Apache poi\\test.xlsx");
// take file in read mode
		FileInputStream fis = new FileInputStream(f);
// take control of entire workbook
		Workbook wb = WorkbookFactory.create(fis);
// take control of sheet
		Sheet sh = wb.getSheet("Sheet1");
//take control of row
		//Row r = sh.getRow(2);
		Row r = sh.createRow(6);
// take control of cell
	//	Cell c = r.getCell(0);
		Cell c = r.createCell(0);
c.setCellValue("usr6");
r.createCell(1).setCellValue("ijg768");
//r.getCell(1).setCellValue("bvb765");		
	FileOutputStream fos = new FileOutputStream(f);	
		wb.write(fos);
		// close the workbook
		wb.close();
		System.out.println("Data written successfully");
	}
}
