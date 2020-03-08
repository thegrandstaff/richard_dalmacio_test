package test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class login
{
    public static void main(String[] args) throws Exception
    {
        File source = new File("C:\\ExcelData\\COMP316Test.xlsx");

        FileInputStream filesource = new FileInputStream(source);

        XSSFWorkbook wb = new XSSFWorkbook(filesource);

        XSSFSheet sheet1 = wb.getSheetAt(0);

        //write in excel sheet : username
        sheet1.getRow(0).createCell(0).setCellValue("richard");

        //write in excel sheet : password
        sheet1.getRow(0).createCell(1).setCellValue("Password123");

        //read username
        String fakeUsername = sheet1.getRow(0).getCell(0).getStringCellValue();

        //read password
        String fakePassword = sheet1.getRow(0).getCell(1).getStringCellValue();

        //path to access chromedriver
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        //goes to website with login
        driver.get("http://demo.guru99.com/test/login.html");

        //get email webelement
        WebElement email = driver.findElement(By.id("email"));

        //get password webelement
        WebElement password = driver.findElement(By.name("passwd"));

        //insert email
        email.sendKeys(fakeUsername);

        //insert password
        password.sendKeys(fakePassword);

        //this is just to take a break while I take a screenshot
        Thread.sleep(3000);

        //find login button
        WebElement login = driver.findElement(By.id("SubmitLogin"));
        //click login button
        login.submit();

        try {
            //whenever you want to write, you have to use FileOutputStream
            FileOutputStream fileOut = new FileOutputStream(source);
            wb.write(fileOut);
        } catch (Exception e) {
            // catch block
            e.printStackTrace();
        }

        wb.close();
    }

}
