package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelDemo {
	
	static XSSFRow row;

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		//blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		
		//prepare data to write 
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"Id", "NAME", "LASTNAME"});
		data.put("2", new Object[] {1, "Somu", "Kummara"});
		data.put("3", new Object[] {2, "Sujatha", "Kummara"});
		data.put("4", new Object[] {3, "Charanya", "Kummara"});
		data.put("5", new Object[] {4, "CharanTeja", "Kummara"});
		
		
		
		//iterate over data write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			
			XSSFRow row = sheet.createRow(rownum++);
			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
				cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}
			
		}
		try {
			FileOutputStream out= new FileOutputStream(new String("POI//Howtodojava_demo.xlsx"));
					workbook.write(out);
					out.close();
					System.out.println("Howtodojava_demo.xlsx written successfully");
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
}


