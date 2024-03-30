package Selenium.UploadAndDownloadFile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WriteExcel {

	@Test
	public void Check() throws IOException, InterruptedException {

		String fruit = "Mango";
		
		String value = "1000";
		
		String path = "C://Users//851101//Downloads//download.xlsx";
		
		WebDriverManager.chromedriver().setup();

		WebDriver driver = new ChromeDriver();

		driver.manage().window().maximize();

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		// Download

		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

		 driver.findElement(By.cssSelector("#downloadButton")).click();
		 
		 Thread.sleep(2000);

		// GetcolumnNumber
		int columnNumber = getColumnNumber(path, "Price");

		System.out.println(columnNumber);
		// Getrownumber
		int rowNumber = getRowNumber(path, "Mango");

		System.out.println(rowNumber);

		// Updatecell
		updateCell(path, rowNumber, columnNumber, value);

		// Upload

		WebElement upload = driver.findElement(By.cssSelector("#fileinput"));

		upload.sendKeys("C://Users//851101//Downloads//download.xlsx");

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector(".Toastify__toast-icon")));

		String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");

		String actualPrice = driver.findElement(By.xpath(
				"//div[text()='" + fruit + "']/parent::div/parent::div/div[@id='cell-" + priceColumn + "-undefined']"))
				.getText();

		System.out.println(actualPrice);

		Assert.assertEquals(value, actualPrice);
		
		Thread.sleep(10000);

		driver.close();

		// Toastify__toast-icon Toastify--animate-icon Toastify__zoom-enter
		// //div[text()='Mango']/parent::div/parent::div/div
	}

	private void updateCell(String path, int rowNumber, int columnNumber, String value) throws IOException {
		// TODO Auto-generated method stub

		XSSFSheet sheet = null;

		Row row1;

		FileInputStream fis = new FileInputStream(path);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {

				sheet = workbook.getSheetAt(i);

				Row row = sheet.getRow(rowNumber);

				Cell cell = row.getCell(columnNumber);

				cell.setCellValue(value);
				
				FileOutputStream fos = new FileOutputStream(path);
				
				workbook.write(fos);
				
				workbook.close();
				
				fis.close();
				
				fos.close();
				
				break;
			}

		}
	}

	private int getColumnNumber(String path, String price) throws IOException {
		// TODO Auto-generated method stub

		XSSFSheet sheet = null;

		int columncount = 0;

		int k = 0;

		FileInputStream file = new FileInputStream(path);

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {

				sheet = workbook.getSheetAt(i);

				break;

			}
		}
		Iterator<Row> row = sheet.rowIterator();

		Row firstrow = row.next();

		Iterator<Cell> cells = firstrow.cellIterator();

		while (cells.hasNext()) {
			Cell value = cells.next();

			if (value.getStringCellValue().equalsIgnoreCase(price)) {

				columncount = k;
			}
			k++;
		}

		return columncount;
	}

	private int getRowNumber(String path, String fruit) throws IOException {
		// TODO Auto-generated method stub
		XSSFSheet sheet = null;

		int rowcount = 0;
		int columncount = 0;

		int k = 0;

		int m = 0;

		FileInputStream file = new FileInputStream(path);

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {

				sheet = workbook.getSheetAt(i);

			}
		}

		String name = "fruit";

		Iterator<Row> row1 = sheet.rowIterator();

		while (row1.hasNext()) {

			Row cuurentrow = row1.next();

			Iterator<Cell> cells1 = cuurentrow.cellIterator();

			while (cells1.hasNext()) {
				Cell value = cells1.next();

				if (value.getCellType() == CellType.STRING) {

					if (value.getStringCellValue().equalsIgnoreCase(fruit)) {
						name = value.getStringCellValue().toString();

						System.out.println(name);

						break;
					}
				} else if (value.getCellType() == CellType.NUMERIC) {
					name = NumberToTextConverter.toText(value.getNumericCellValue()).toString();
				}

			}
			rowcount = m;
			m++;
			if (name.equalsIgnoreCase(fruit)) {
				break;
			}
		}

		return rowcount;
	}
}
