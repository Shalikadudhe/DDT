package com.dprovider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.Test;

import com.config.Excel_Reader;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DPdemo extends Excel_Reader {
	public static void writeResult(String fileNm, String sheetNm, String result)
			throws EncryptedDocumentException, IOException {
		try {
			FileInputStream fis = new FileInputStream(fileNm);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet sh = wb.getSheet(sheetNm);
			int rows = sh.getLastRowNum();
			for (int j = 1; j <rows; j++) {
				Row r = sh.getRow(j);
				Cell cc = r.createCell(3);
				cc.setCellValue(result);
				break;
			}

			FileOutputStream fos = new FileOutputStream(fileNm);
			wb.write(fos);
			fos.close();

					} catch (FileNotFoundException e) {
						e.printStackTrace();
					}
	}

	@Test(dataProvider = "DDT_TEST", dataProviderClass = DataProviderC.class)
	public static void testUserNmAndPassword(String userNm, String password) throws InterruptedException, IOException {
		// System.out.println("userNm " + userNm + " password " + password);
		WebDriverManager.chromedriver().setup();
		RemoteWebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2));
		driver.get("https://demowebshop.tricentis.com/login");
		driver.findElement(By.id("Email")).sendKeys(userNm);
		driver.findElement(By.id("Password")).sendKeys(password);
		driver.findElement(By.xpath("//input[@value='Log in']")).click();

		String result = null;

		try {
			Boolean isloggedIn = driver.findElement(By.xpath("//a[@href='/logout']")).isDisplayed();
			if (isloggedIn == true) {
				result = "PASS";
			}

			System.out.println(
					"Usernm : " + userNm + "=====> Password :" + password + "====> Login Successfull ?====>" + result);

			writeResult("E:\\SELENIUM-Workspace\\DDT_CrossBrowser_Screenshot\\src\\test_Data\\DataDrivenFile.xlsx",
					"sheet1", result);
			driver.findElement(By.xpath("//a[@href='/logout']")).click();
		} catch (Exception e) {

			Boolean ErrorIs = driver.findElement(By.xpath("//*[@class='inputs']/span/span")).isDisplayed();

			if (ErrorIs == true) {
				result = "FAIL";
			}
			System.out.println(
					"Usernm : " + userNm + "=====> Password :" + password + "====> Login Successfull ?====>" + result);
			writeResult("E:\\SELENIUM-Workspace\\DDT_CrossBrowser_Screenshot\\src\\test_Data\\DataDrivenFile.xlsx",
					"sheet1", result);
			// driver.findElement(By.xpath("//a[@href='/logout']")).click();
			driver.findElement(By.xpath("//a[@href='/login']")).click();
		}

		Thread.sleep(1000);
		driver.findElement(By.xpath("//a[@href='/login']")).click();
		/*
		 * FileOutputStream fos = new FileOutputStream(
		 * "E:\\SELENIUM-Workspace\\DDT_CrossBrowser_Screenshot\\src\\test_Data\\DDT.xlsx"
		 * ); wb.write(fos); wb.close();
		 */
		driver.close();

	}
}
