package com.cydeo.tests;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelRead {

    @Test
            public void read_from_excel_file() throws IOException {
        String path="SampleData.xlsx";

        //to read from excell we need to loaad it FileInput Stream
        FileInputStream fileInputStream=new FileInputStream(path);

        //workbook>sheet>row>cell


        // <1>  Create aa workbook
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);

        //<2> We need to get specific sheet from currently opened workbook
        XSSFSheet sheet=workbook.getSheet("Employees");

        //<3> Select row and cell
        //Print out mary's cell
        //Indexes start from 0
        System.out.println(sheet.getRow(1).getCell(0));

        //print out developer
        System.out.println(sheet.getRow(3).getCell(2));


        //return the count of used cells only
        //starts from counting from 1

        int usedRows=sheet.getPhysicalNumberOfRows();
        System.out.println("usedRows = " + usedRows);

        //Returns the number from top cell to bottom cell
        //It doesn't care if the cells is empty or not
        //starts conting from 0

        int lastUsed=sheet.getLastRowNum();
        System.out.println("lastUsed = " + lastUsed);

        //TODO: Create a logic to print Vinod's name
        for (int rowNum=0; rowNum<usedRows; rowNum++){
            if(sheet.getRow(rowNum).getCell(0).toString().equals("Vinod")){
                System.out.println(sheet.getRow(rowNum).getCell(0));
            }
        }

        //TODO: Create a logic to print out Linda's Job_ID
        //Check if name is Linda --> print out JOB_ID of Linda
        for(int rowNum=0; rowNum<usedRows; rowNum++){
            if(sheet.getRow(rowNum).getCell(0).toString().equals("Linda")){
                System.out.println("Linda's job ID: "+ sheet.getRow(rowNum).getCell(2));
            }
        }
    }








}
