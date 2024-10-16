/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package org.visu.demo;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

/**
 *
 * @author SHREYA
 */
public class SeleniumDemo {

    public static void main(String[] args) {

        try {
            XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\SHREYA\\project\\file_example_XLSX_10.xlsx");
            WebDriver driver = new ChromeDriver();
            System.out.println("in...");
            driver.navigate().to("https://mail.office365.com/");
            Thread.sleep(5000);
            WebElement we = driver.findElement(By.id("i0116"));
            we.sendKeys("harshad");


            XSSFSheet sheet = book.getSheetAt(0);

            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    // Check the cell type and format accordingly
                    System.out.println(cell.getColumnIndex());
                    if (cell.getColumnIndex() == 1) {

                    }
                    // switch (cell.getCellType())
                    // {
                    // // case CellType.NUMERIC -> System.out.print(cell.getNumericCellValue() +
                    // "\t");
                    // // case CellType.STRING -> System.out.print(cell.getStringCellValue() +
                    // "\t");
                    // // default -> System.err.println(".");
                    // }
                }
                book.close();
                System.out.println("");

            }

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (InterruptedException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
