package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HandsOnPOI1 {
	static File f;
	static FileOutputStream fos;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFRow row1,row2,row3,row4,row5,row6,row7,row8,row9;
	static XSSFCell cell;



	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub//write-output,read-input
		f=new File(System.getProperty("user.dir")+"//POI//HandsOn1.xlsx");
		fos=new FileOutputStream(f);
		wb=new XSSFWorkbook();
		sheet=wb.createSheet("Details");
		row=sheet.createRow(0);
		cell=row.createCell(0);
		cell.setCellValue("NithishKumar");
		row1=sheet.createRow(1);
		cell=row1.createCell(0);
		cell.setCellValue("vagicharlanithish@gmail.com");
		row2=sheet.createRow(2);
		cell=row2.createCell(0);
		cell.setCellValue("8328538984");
		row3=sheet.createRow(3);
		cell=row3.createCell(0);
		cell.setCellValue("Rollno:42");
		row4=sheet.createRow(4);
		cell=row4.createCell(0);
		cell.setCellValue("Nithish@1004");
		row5=sheet.createRow(5);
		cell=row5.createCell(0);
		cell.setCellValue("Cricket");
		row6=sheet.createRow(6);
		cell=row6.createCell(0);
		cell.setCellValue("Male");
		row7=sheet.createRow(7);
		cell=row7.createCell(0);
		cell.setCellValue("OldTown");
		row8=sheet.createRow(8);
		cell=row8.createCell(0);
		cell.setCellValue("Markapur");
		row9=sheet.createRow(9);
		cell=row9.createCell(0);
		cell.setCellValue("Prakasham");
		
		int noOfRows=sheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows::"+noOfRows);
		
		int noOfCells=row1.getPhysicalNumberOfCells();
		System.out.println("Number of Cells::"+noOfCells);
		
		
			for(int j=0; j<noOfCells; j++)
			{
				System.out.println(sheet.getRow(0).getCell(j).getStringCellValue());
			}

		
		/*fos=new FileOutputStream(f);
		try {
			wb.write(fos);
			System.out.println("Personal Details");
			fos.close();
			wb.close();
			
		} catch (IOException e) {
			e.printStackTrace();

	}*/

}
}
