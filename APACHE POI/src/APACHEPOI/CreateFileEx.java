package APACHEPOI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class CreateFileEx {

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		
		File f1=new File(System.getProperty("user.dir")+"//POI//Employee.xls");
		FileOutputStream fos=new FileOutputStream (f1);
		System.out.println("File is Successfully Created");

	}

}
