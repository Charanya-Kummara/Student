package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HandsOnPOI {
	static File f;
	static HSSFWorkbook wb;
	static HSSFSheet sheet;
	static HSSFRow row;
	static HSSFCell cell;
	static HSSFCell cell1;
	static HSSFCell cell2,cell3,cell4,cell5,cell6,cell7,cell8,cell9;
	static FileOutputStream fos;

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		 f=new File(System.getProperty("user.dir")+"//POI//HandsOn.xls");
			
			/**
			 * 
			 * 
			 */
			wb=new HSSFWorkbook();
			sheet=wb.createSheet("Details");
			
			 row=sheet.createRow(0);
			 cell=row.createCell(0);
			 cell1=row.createCell(1);
			 cell2=row.createCell(2);
			 cell3=row.createCell(3);
			 cell4=row.createCell(4);
			 cell5=row.createCell(5);
			 cell6=row.createCell(6);
			 cell7=row.createCell(7);
			 cell8=row.createCell(8);
			 cell9=row.createCell(9);
			
			//set input in cell0 and row0
			cell.setCellValue("Charanya");
			cell1.setCellValue("Charanyakummara@gmail.com");
			cell2.setCellValue("9014798545");
			cell3.setCellValue("Rollno:30");
			cell4.setCellValue("Cherry@1004");
			cell5.setCellValue("Cricket");
			cell6.setCellValue("Female");
			cell7.setCellValue("KovurNagar");
			cell8.setCellValue("Anantapur");
			cell9.setCellValue("AndhraPradesh");
			
			fos=new FileOutputStream(f);
			try {
			wb.write(fos);
			System.out.println("Personal Details are:");
			} catch (IOException e) {
				e.printStackTrace();
			}

	}

}
