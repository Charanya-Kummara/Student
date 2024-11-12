package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ReadExcel {
	static File f1;
	static FileInputStream fis;
	static HSSFWorkbook wb1;
	static HSSFSheet sheet1;
	static HSSFRow row1;
	static HSSFCell cell0;
	static HSSFCell cell10;
	static HSSFCell cell20;

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/**
		 * read data from excel
		 * display all data from excel file in console
		 * ===============================================================
		 * File
		 * FileInputStream
		 * =================================================================================================
		 * Get a workbook ------>HSSFworkbook
		 * Get a sheet    ------>HSSFsheet
		 * Get a row      ------>HSSFrow
		 * Get a cell     ------>HSSFcell
		 * 
		 */
		f1=new File(System.getProperty("user.dir")+"//POI//Employee.xls");
		fis=new FileInputStream(f1);
		wb1=new HSSFWorkbook(fis);
		sheet1=wb1.getSheet("LoginCredentials");
		row1=sheet1.getRow(0);
		cell0=row1.getCell(0);
		//String cell01=String.valueOf(cell0);
		//System.out.println(cell01);
		cell10=row1.getCell(1);
		cell20=row1.getCell(2);
		//System.out.println(String.valueOf(cell0.getNumericCellValue()));
		//System.out.println(cell10.getStringCellValue());
		//System.out.println(cell20.getStringCellValue());
		int noOfRows=sheet1.getPhysicalNumberOfRows();
		System.out.println("Number of Rows::"+noOfRows);
		
		int noOfCells=row1.getPhysicalNumberOfCells();
		System.out.println("Number of Cells::"+noOfCells);
		
		/*int i=0;
		while(i<noOfCells)
		{
			System.out.println(sheet1.getRow(0).getCell(i).getRichStringCellValue());
			i++;
		}*/
		
			for(int j=0; j<noOfCells; j++)
			{
				System.out.println(sheet1.getRow(0).getCell(j).getNumericCellValue());
			}
		}
		
	}


