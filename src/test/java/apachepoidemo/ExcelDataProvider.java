package apachepoidemo;

import org.testng.annotations.AfterMethod;
import org.testng.annotations.Test;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

public class ExcelDataProvider {
	public XSSFWorkbook wb;
	public XSSFSheet ws;
	public XSSFRow row;
	public XSSFCell cell;
	public int rowcount;
	public FileInputStream fis;
	public File src;

	@BeforeSuite
	public void setup() throws IOException {
		src = new File("./linkedin.xlsx");
		//src = new File("/home/swapnil/ECLIPSE/eclipse-workspace/apachepoidemo/linkedin.xlsx");
		fis = new FileInputStream(src);
		wb = new XSSFWorkbook(fis);
		ws = wb.getSheetAt(0);
		rowcount = ws.getLastRowNum();
	}

	@Test(priority = 0)
	public void checkFileExistance() {
		boolean status = src.exists();

		try {
			if (status == true) {
				System.out.println("Excel File is found, You can execute test cases");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// TC= 1 : to get data from single cell
	@Test(priority = 1, dependsOnMethods = "checkFileExistance")
	public void getDatafromSingleCell() {
		System.out.println("$$ TC01 started $$");
		// step1: define the starting row position
		row = ws.getRow(0);
		cell = row.getCell(0);
		System.out.println(ws.getRow(0).getCell(0).getStringCellValue());
	}

	// TC= 2 : to get data from single row multiple cols
	@Test(priority = 2, dependsOnMethods = "checkFileExistance")
	public void getDatafromSingleRowMultipleCols() {
		System.out.println("## TC02 started ##");
		// step1: define the starting row position
		row = ws.getRow(1);
		int cellcount = row.getLastCellNum();
		for (int i = 0; i < cellcount; i++) {
			System.out.println(ws.getRow(1).getCell(i).getStringCellValue());
		}

	}

	// TC= 3 : to get data from multiple rows single cols
	@Test(priority = 3, dependsOnMethods = "checkFileExistance")
	public void getDatafromMultipleRowsSingleCol() {
		System.out.println("** TC03 started **");
		// step1: define the starting row position
		for (int i = 0; i <= rowcount; i++) {
			System.out.println(ws.getRow(i).getCell(2).getStringCellValue());
		}

	}

	// TC= 4 : to get data from multiple rows and multiple cols
	@Test(priority = 4, dependsOnMethods = "checkFileExistance")
	public void getDatafromMultipleRowsMultipleCols() {
		System.out.println("%% TC04 started %%");
		// step1: define the starting row position
		for (int i = 0; i <= rowcount; i++) {
			row = ws.getRow(i);
			int colcount = row.getLastCellNum();

			for (int j = 0; j < colcount; j++) {
				System.out.println(ws.getRow(i).getCell(j).getStringCellValue());
			}
		}
	}

	
	@AfterSuite
	public void tearDown() throws IOException {
		fis.close();
	}

}
