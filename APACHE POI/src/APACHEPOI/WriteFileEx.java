package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteFileEx {

	static File f1;
	static HSSFWorkbook wb1;
	static HSSFSheet sheet1;
	static HSSFRow row1;
	static HSSFCell cell0;
	static HSSFCell cell10;
	static HSSFCell cell20;
	static FileOutputStream fos1;
	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		
		f1=new File(System.getProperty("user.dir")+"//POI//Employee.xls");
		
		/**
		 * 
		 * 
		 */
		wb1=new HSSFWorkbook();
		
		sheet1=wb1.createSheet("LoginCredentials");
		//System.out.println("Name of the Sheet::"+sheet.getSheetName());
		 row1=sheet1.createRow(0);
		
		 cell0=row1.createCell(0);
		 cell10=row1.createCell(1);
		 cell20=row1.createCell(2);
		
		//set input in cell0 and row0
		cell0.setCellValue("101");
		cell10.setCellValue("Charanya");
		cell20.setCellValue("Anantapur");
		
		fos1=new FileOutputStream(f1);
		//System.out.println("File is Successfully Created");
		try {
		wb1.write(fos1);
		System.out.println("Data is written Successfully");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	}


