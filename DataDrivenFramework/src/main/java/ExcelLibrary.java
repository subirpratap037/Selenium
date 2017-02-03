import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

/**
 * Created by 3Embed on 2/3/2017.
 */
public class ExcelLibrary
{
    public String getExcelData(String SheetName, int rowNum, int cellNum) {
        String retvalue = null;
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\3Embed\\Desktop\\Login.xlsx");
            Workbook wb = WorkbookFactory.create(fis);
            Sheet s = wb.getSheet(SheetName);
            Row r = s.getRow(rowNum);
            Cell c = r.getCell(cellNum);
            retvalue = c.getStringCellValue();


        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return retvalue;
    }

    public int getLastRowCount(String SheetName){
        int rowcount =0;
        try{
            FileInputStream fis = new FileInputStream("C:\\Users\\3Embed\\Desktop\\Login.xlsx");
            Workbook wb = WorkbookFactory.create(fis);
            Sheet s = wb.getSheet(SheetName);
            rowcount = s.getLastRowNum();

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return rowcount;
    }

    public String writeDatatoExcel(String Sheet, int rowNum, int cellNum, String data){
        try{
            FileInputStream fis = new FileInputStream("C:\\Users\\3Embed\\Desktop\\Login.xlsx");
            Workbook wb = WorkbookFactory.create(fis);
            Sheet s = wb.getSheet(Sheet);
            Row r = s.getRow(rowNum);
            Cell c = r.createCell(cellNum);
            c.setCellValue(data);

            FileOutputStream fos = new FileOutputStream("C:\\Users\\3Embed\\Desktop\\Login.xlsx");
            wb.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    Class LoginLogoutDriven(){
        public static void main(String args[]){
            ExcelLibrary xlib= new ExcelLibrary();

            int rowcount = xlib.getLastRowCount("Login");
            System.out.println(rowcount);

            WebDriver driver = new FirefoxDriver();
            driver.get("https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1#identifier");
            driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
            for(int i=1; i<=rowcount; i++){
                String un = xlib.getExcelData("Login", i, 0);
                String pw = xlib.getExcelData("Login", i, 1);
                driver.findElement(By.id("Email")).sendKeys("3embedsoft@gmail.com");
                driver.findElement(By.id("Passwd")).sendKeys("3embed007");
                driver.findElement(By.id("signIn")).click();
                try {
                    driver.findElement(By.className("gb_9a gbii")).click();
                    driver.findElement(By.id("gb_71")).click();
                    xlib.writeDatatoExcel("Login",i,2,"pass");
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                    xlib.writeDatatoExcel("Login",i,2,"fail");
                }
            }

        }
    }
}

