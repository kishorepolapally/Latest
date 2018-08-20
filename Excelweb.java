package SGB;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.ie.InternetExplorerDriver;
//import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.Test;


public class Excelweb {
	
	@Test
	public void f() throws Exception{

		System.setProperty("webdriver.chrome.driver", "browser/chromedriver.exe");
		WebDriver driver = new ChromeDriver();   
		driver.get("https://www.w3schools.com/html/html_tables.asp");
		driver.manage().window().maximize();
		
		String FilePath = "E:\\kishore.xlsx";
		String SheetName = "Sheet1";
		FileInputStream fi = new FileInputStream(FilePath);  //Path
		XSSFWorkbook wb = new XSSFWorkbook(fi);  //xlsx format
		XSSFSheet sh = wb.getSheet(SheetName); // sheet name
		int rowCount = sh.getLastRowNum()+1;
		System.out.println(rowCount+"123456");
		for (int i = 1; i <rowCount; i++) {
		XSSFRow r = sh.createRow(i);     // Refernece here 0 indicate the row number
		String k =ExcelRead(i,0).trim();
		
		if (!k.equals(null) || k.equals("")) {
			String value = driver.findElement(By.xpath("//table[@id='customers']//td[text()='"+k+"']/../td[3]")).getText();
			Cell c = r.createCell(3);  // here 0 indicate the column. COlumn always start with O
			c.setCellValue(value);
		}
		else {
			System.out.println("invalid data in cell");
		}
	}
		FileOutputStream fo = new FileOutputStream(FilePath);
		wb.write(fo);
		fo.close();
		wb.close();
	}
	

	public String ExcelRead(int row, int Column) throws Exception{
		DataFormatter d = new DataFormatter();
		String FilePath = "E:\\kishore.xlsx";
		String SheetName = "Sheet1";
		String bal = "";
		FileInputStream fi = new FileInputStream(FilePath);  //Path
		XSSFWorkbook wb = new XSSFWorkbook(fi);  //xlsx format
		XSSFSheet sh = wb.getSheet(SheetName); // sheet name
		int rowcount = sh.getRow(0).getLastCellNum();
		XSSFRow r = sh.getRow(row);  //row
		String cellvalue = d.formatCellValue(r.getCell(Column));  // column
//		System.out.println("cellvalue"+cellvalue);
		wb.close();
		return cellvalue;
		
	}


	    public String Excelwrite(int row, int Column, String data) throws Exception{
		DataFormatter d = new DataFormatter();
		String FilePath = "E:\\kishore.xlsx";
		String SheetName = "Sheet1";
		String bal = "";
		FileInputStream fi = new FileInputStream(FilePath);  //Path
		XSSFWorkbook wb = new XSSFWorkbook(fi);  //xlsx format
		XSSFSheet sh = wb.getSheet(SheetName); // sheet name
		int rowcount = sh.getRow(1).getLastCellNum();
		XSSFRow r = sh.getRow(row);  //row
		String cellvalue = d.formatCellValue(r.getCell(Column));  // column
		System.out.println("cellvalue"+cellvalue);
		FileOutputStream fo = new FileOutputStream(FilePath);
		XSSFRow ro = sh.createRow(row);     // Reference here 0 indicate the row number
		Cell c = ro.createCell(Column);  // here 0 indicate the column. COlumn always start with O
		c.setCellValue(data);
		wb.write(fo);
		fo.close();
		wb.close();
		return cellvalue;
		
	}

	
	public int rowcount() throws Exception{
		DataFormatter d = new DataFormatter();
		String FilePath = "E:\\kishore.xlsx";
		String SheetName = "Sheet1";
		FileInputStream fi = new FileInputStream(FilePath);  //Path
		XSSFWorkbook wb = new XSSFWorkbook(fi);  //xlsx format
		XSSFSheet sh = wb.getSheet(SheetName); // sheet name
//		int rowcount = sh.getRow(0).getLastCellNum();  //columns
		int rowCount = sh.getLastRowNum()+1;
		return rowCount;
		
		
	}
	
	
	
	
}