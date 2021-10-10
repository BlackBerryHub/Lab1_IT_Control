package com.company;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.opera.OperaDriver;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {

        Scanner in = new Scanner(System.in);
        System.out.print("Input institute (Example:ІКНІ): ");
        String institute = in.nextLine();

        System.out.print("Input group (Example:КН-405): ");
        String group = in.nextLine();

        System.setProperty("webdriver.opera.driver", "C://Users//marat//Desktop//selenium driver//drivers//operadriver_win64/operadriver.exe");

        OperaDriver driver=new OperaDriver();


        driver.get("https://student.lpnu.ua/students_schedule?departmentparent_abbrname_selective="+institute+"&studygroup_abbrname_selective="+group+"&semestrduration=1");

        Workbook wb = new HSSFWorkbook();
        Sheet s = wb.createSheet("Schedule "+institute+" "+group);

        Row row =  s.createRow(0);
        row.createCell(0).setCellValue("Пн");
        row.createCell(1).setCellValue("Вт");
        row.createCell(2).setCellValue("Ср");
        row.createCell(3).setCellValue("Чт");
        row.createCell(4).setCellValue("Пт");

        List<WebElement> contentList = driver.findElements(By.xpath("//div[@class='view-content']/div"));
        int i = 0;
        int j = 1;
        for(WebElement element:contentList){
            System.out.println("Section "+i+":");
            List<WebElement> dayList = contentList.get(i).findElements(By.cssSelector("div.stud_schedule"));
            for(WebElement element1:dayList){
                System.out.println(element1.getText());
                Row row1 = s.createRow(j);
                row1.createCell(i).setCellValue(element1.getText());
                j++;
            }
            i++;
        }

        try {
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\marat\\Desktop\\Schedule.xls");
            wb.write(fileOut);
            fileOut.close();
        }
        catch (Exception e){
        }
    }
}
