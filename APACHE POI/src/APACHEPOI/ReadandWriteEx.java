package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadandWriteEx {
	static File f;
	static FileOutputStream fos;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFRow row1;
	static XSSFRow row2;
	static XSSFRow row3;
	static XSSFRow row4;
	static XSSFRow row5;
	static XSSFCell cell;
	static XSSFCell cell1;
	static XSSFCell cell2;

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		f=new File(System.getProperty("user.dir")+"//POI//Languages.xlsx");
		wb=new XSSFWorkbook();
		sheet=wb.createSheet("List of Programming Languages");
		row=sheet.createRow(0);
		cell=row.createCell(0);
		cell.setCellValue("Java");
		row1=sheet.createRow(1);
		cell=row1.createCell(0);
		cell.setCellValue("Python");
		row2=sheet.createRow(2);
		cell=row2.createCell(0);
		cell.setCellValue("C#");
		row3=sheet.createRow(3);
		cell=row3.createCell(0);
		cell.setCellValue("Php");
		row4=sheet.createRow(4);
		cell=row4.createCell(0);
		cell.setCellValue("Ruby");
		row5=sheet.createRow(5);
		cell=row5.createCell(0);
		cell.setCellValue("VB");
		
		fos=new FileOutputStream(f);
		try {
			wb.write(fos);
			System.out.println("Data mentioned Below!");
			fos.close();
			wb.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		

	}

}
