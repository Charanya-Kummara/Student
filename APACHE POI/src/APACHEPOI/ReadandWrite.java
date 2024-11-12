package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileOutputStream;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadandWrite {
	static File f;
	static FileInputStream fis;
	static FileOutputStream fos;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFRow row1;
	static XSSFRow row2;
	static XSSFCell cell;
	static XSSFCell cell1,cell2,cell3;
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/**
		 * XSSF API
		 * perform reading and writing operation over
		 * the excel file which is having format .xlsx
		 * ======================================================
		 * operations
		 * create an excel .xlsx file
		 * write data in the file
		 * display or read data from the created excel file
		 * =====================================================
		 * File
		 * FileOutputStream
		 * =====================================================
		 *  create a workbook ------>XSSFworkbook
		 * create a sheet    ------>XSSFsheet
		 * create a row      ------>XSSFrow
		 * create a cell     ------>XSSFcell
		 * 
		 * 
		 * 
		 * 
		 * 
		 */
		
		
		f=new File(System.getProperty("user.dir")+"//POI//Department.xlsx");
		fis=new FileInputStream(f);
		wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet("DepartmentNames");
		row=sheet.getRow(0);
		cell=row.getCell(0);
		row1=sheet.getRow(1);
		cell1=row1.getCell(0);
		row2=sheet.getRow(2);
		cell2=row2.getCell(0);
		
		
		System.out.println(cell.getStringCellValue());
		System.out.println(cell1.getStringCellValue());
		System.out.println(cell2.getStringCellValue());
		
		
		
		/*sheet=wb.createSheet("DepartmentNames");
		row=sheet.createRow(0);
		cell=row.createCell(0);
		cell.setCellValue("IT Department");
		row1=sheet.createRow(1);
		cell=row1.createCell(0);
		cell.setCellValue("Finance");
		row2=sheet.createRow(2);
		cell=row2.createCell(0);
		cell.setCellValue("HR");
		row3=sheet.createRow(3);
		cell=row3.createCell(0);
		cell.setCellValue("Marketing");
		
		fos=new FileOutputStream(f);
		try {
			wb.write(fos);
			System.out.println("Data is written in Sheet!");
			fos.close();
			wb.close();
			
		} catch (IOException e) {
			e.printStackTrace();*/
		}
		//System.out.println("File is Successfully Created");
	}


