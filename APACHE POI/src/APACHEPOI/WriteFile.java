package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteFile {
	static File f;
	static HSSFWorkbook wb;
	static HSSFSheet sheet;
	static HSSFRow row;
	static HSSFCell cell;
	static HSSFCell cell1;
	static HSSFCell cell2;
	static FileOutputStream fos;

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		
		/**
		 * HSSF API
		 * ==========================
		 * create a workbook ------>HSSFworkbook
		 * create a sheet    ------>HSSFsheet
		 * create a row      ------>HSSFrow
		 * create a cell     ------>HSSFcell
		 */
		 f=new File(System.getProperty("user.dir")+"//POI//ApachePOI.xls");
		
		/**
		 * 
		 * 
		 */
		wb=new HSSFWorkbook();
		
		sheet=wb.createSheet("LoginCredentials");
		//System.out.println("Name of the Sheet::"+sheet.getSheetName());
		 row=sheet.createRow(0);
		
		 cell=row.createCell(0);
		 cell1=row.createCell(1);
		 cell2=row.createCell(2);
		
		//set input in cell0 and row0
		cell.setCellValue("Charanyakummara@gmail.com");
		cell1.setCellValue("6879862");
		cell2.setCellValue("Bengalore");
		
		fos=new FileOutputStream(f);
		//System.out.println("File is Successfully Created");
		try {
		wb.write(fos);
		System.out.println("Data is written Successfully");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
