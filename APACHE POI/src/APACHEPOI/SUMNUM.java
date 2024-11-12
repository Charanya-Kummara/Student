package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SUMNUM {
	static File f;
	static FileInputStream fis;
	static FileOutputStream fos;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFCell cell1,cell2;
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		f=new File(System.getProperty("user.dir")+"//POI//Addition.xlsx");
		fos=new FileOutputStream(f);
		wb=new XSSFWorkbook(fos);

		            // Access the first sheet
		sheet=wb.getSheet("SumOfTwoNum");
		            
		           
		            row = sheet.getRow(0); // First row
		             cell1 = row.getCell(0); // First cell (A1)
		             cell2 = row.getCell(1); // Second cell (B1)

		            // Read the numeric values
		            double number1 = cell1.getNumericCellValue();
		            double number2 = cell2.getNumericCellValue();

		            // Sum the numbers
		            double sum = number1 + number2;

		            // Output 
		            System.out.println("The sum of " + number1 + " and " + number2 + " is: " + sum);

		        
		}


	}


