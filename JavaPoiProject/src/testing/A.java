package testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class A {

	public static void main(String aa[]) throws Exception
	{
		File f = new File("G:\\Data21.xlsx");
		FileOutputStream fo = new FileOutputStream(f);
		XSSFWorkbook wk =  new XSSFWorkbook();
		XSSFSheet s1 = wk.createSheet();
		Row r1= s1.createRow(0);
		Cell c1= r1.createCell(0);
		c1.setCellValue("HELLO DELHI1");
		wk.write(fo);
		fo.flush();
		fo.close();
		System.out.println("Changes Made");
	}
	
	
}
