package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class CreateFile {

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		/* create an excel file and set test inputs
		 * and perform all the regarding operations
		 */
		/*file class
		 * FileOutputStream 
		 */
		
		
		
		
		File f=new File(System.getProperty("user.dir")+"//POI//ApachePOI.xls");
		FileOutputStream fos=new FileOutputStream (f);
		System.out.println("File is Successfully Created");

	}

}
