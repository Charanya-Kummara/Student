package APACHEPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteExcelEx {
	
	static XSSFCell cell1,cell2;
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//blank workbook
				XSSFWorkbook workbook = new XSSFWorkbook();
				
				//create a blank sheet
				XSSFSheet sheet = workbook.createSheet("Department_Company");
		        

				//prepare data to write 
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put("1", new Object[] {"Company", "Department"});
				data.put("2", new Object[] {"Infosys", "IT"});
				data.put("3", new Object[] {"Cognizant", "Solution"});
				data.put("4", new Object[] {"Capgemini", "QA"});
				data.put("5", new Object[] {"MindTree", "Development"});
				data.put("6", new Object[] {"Accenture", "Cloud Based"});
				
			
				//iterate over data write to sheet
				Set<String> keyset = data.keySet();
				int rownum = 0;
				for (String key : keyset) {
					
					XSSFRow row = sheet.createRow(rownum++);
					Object [] objArr = data.get(key);
					int cellnum = 0;
					for (Object obj : objArr)
					{
						
						//to apply background color
						XSSFCellStyle style = workbook.createCellStyle();
						//XSSFFont font = workbook.createFont();
						
						
				       
				        // RANGE foreground but not the font
				        style.setFillBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
				        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				        Cell cell = row.createCell(cellnum++);
				        cell.setCellStyle(style);
				        
				        
						if(obj instanceof String)
						cell.setCellValue((String)obj);
						else if(obj instanceof Integer)
							cell.setCellValue((Integer)obj);
					}
					
				}
				
				try {
					FileOutputStream out= new FileOutputStream(new String("POI//Howtodojava_EX.xlsx"));
							workbook.write(out);
							out.close();
							System.out.println("Howtodojava_EX.xlsx written successfully");
				}
				catch (Exception e) {
					e.printStackTrace();
				}

			}

	}


