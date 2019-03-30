package cn.tjucis.selenium;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
public class Excel {
	private WebDriver driver;
	private String baseUrl;
	public String[] content = new String[3];

	@Before
	public void setUp() throws Exception {
		  String driverPath = "/C:/Users/hyq/Desktop/lab2/geckodriver.exe";
		  System.setProperty("webdriver.gecko.driver", driverPath);
		  driver = new FirefoxDriver();
		  baseUrl = "http://121.193.130.195:8800";
		  driver.manage().timeouts().implicitlyWait(300, TimeUnit.SECONDS);
	  }
	
	@Test
	public void getcontent() throws Exception
	{
		FileInputStream excelFileInputStream = new FileInputStream("C:\\Users\\hyq\\Desktop\\lab2\\»Ìº˛≤‚ ‘√˚µ•.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
		excelFileInputStream.close();
		XSSFSheet sheet = workbook.getSheetAt(0);
		driver.get(baseUrl + "/");
		for(int rowIndex = 2;rowIndex <= sheet.getLastRowNum();rowIndex++)
		{
			XSSFRow row = sheet.getRow(rowIndex);
			if(row == null){
				continue;
			}
			XSSFCell idCell = row.getCell(1);
			Double d = idCell.getNumericCellValue();
			DecimalFormat df = new DecimalFormat("#.##");
			String idvalue = df.format(d);
			String passwordvalue = idvalue.substring(4);
			WebElement id = driver.findElement(By.name("id"));
			id.click();
			id.clear();
			id.sendKeys(idvalue);
			WebElement password = driver.findElement(By.name("password"));
			password.click();
			password.clear();
			password.sendKeys(passwordvalue);
			WebElement login = driver.findElement(By.id("btn_login"));
			login.click();
			content[2] = driver.findElement(By.id("student-git")).getText();
			Assert.assertEquals(row.getCell(3).getStringCellValue(), content[2]);
			System.out.println(rowIndex - 1);
			WebElement logout = driver.findElement(By.id("btn_logout"));
			logout.click();
			WebElement return1 = driver.findElement(By.id("btn_return"));
			return1.click();
		}
		workbook.close(); 
	}
}
