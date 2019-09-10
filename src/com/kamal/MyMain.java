package com.kamal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

//org.apache.poi:poi:4.1.0
//org.apache.poi:poi-ooxml:4.1.0 //For new version
public class MyMain {

    public static void main(String[] args) {
        //readExcelFile();
        writeExcelFile();
    }

    static class Contact {
        String firstName;
        String lastName;
        String email;
        String dateOfBirth;

        Contact(String firstName, String lastName, String email, String dateOfBirth) {
            this.firstName = firstName;
            this.lastName = lastName;
            this.email = email;
            this.dateOfBirth = dateOfBirth;
        }
    }

    //https://medium.com/@ssaurel/generating-microsoft-excel-xlsx-files-in-java-9508d1b521d9
    private static void writeExcelFile() {
        String[] columns = {"First Name", "Last Name", "Email", "Date Of Birth"};
        List<Contact> contacts = new ArrayList<>();

        try {
            contacts.add(new Contact("Mamun", "Islam", "mamun@gmail.com", "17/01/1980"));
            contacts.add(new Contact("Nasir", "Hossain", "nasir@gmail.com", "17/08/1989"));
            contacts.add(new Contact("Fuad", "Islam", "fuad@gmail.com", "17/07/1956"));
            contacts.add(new Contact("Abdur", "Rahman", "abrahman@gmail.com", "17/05/1988"));

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Contacts");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 14);
            headerFont.setColor(IndexedColors.RED.getIndex());

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // Create a Row
            Row headerRow = sheet.createRow(0);

            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerCellStyle);
            }

            // Create Other rows and cells with contacts data
            int rowNum = 1;

            for (Contact contact : contacts) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(contact.firstName);
                row.createCell(1).setCellValue(contact.lastName);
                row.createCell(2).setCellValue(contact.email);
                row.createCell(3).setCellValue(contact.dateOfBirth);
            }

            // Resize all columns to fit the content size
            for (int i = 0; i < columns.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Kamal\\Downloads\\Book1.xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    //https://medium.com/@ssaurel/reading-microsoft-excel-xlsx-files-in-java-2172f5aaccbe
    private static void readExcelFile() {
        try {
            File mFile = new File("C:\\Users\\Kamal\\Downloads\\Book1.xlsx");
            FileInputStream mInputStream = new FileInputStream(mFile);
            XSSFWorkbook mWorkbook = new XSSFWorkbook(mInputStream);
            XSSFSheet sheet = mWorkbook.getSheetAt(0); // Get first sheet
            // we iterate on rows
            for (Row row : sheet) {
                // iterate on cells for the current row
                Iterator<Cell> mIterator = row.cellIterator();
                while (mIterator.hasNext()) {
                    Cell cell = mIterator.next();
                    System.out.print(cell.toString() + ";");
                }
                System.out.println();
            }

            mWorkbook.close();
            mInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
