package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws FileNotFoundException {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter the absolute path of your file: ");
        String path = scanner.nextLine();

        System.out.print("Enter a word/sentence: ");
        String word = scanner.nextLine();

        try (FileInputStream fis = new FileInputStream(new File(path))) {
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                XSSFSheet sheet = workbook.getSheetAt(i);

                for (Row row : sheet) {

                    for (Cell cell : row) {

                        String cellContent = "";

                        switch (cell.getCellType()) {
                            case STRING:
                                cellContent = cell.getStringCellValue();
                                break;

                            case NUMERIC:

                                if (DateUtil.isCellDateFormatted(cell)) {
                                    cellContent = cell.getDateCellValue().toString();
                                } else {
                                    cellContent = String.valueOf(cell.getNumericCellValue());
                                }
                                break;

                            case BOOLEAN:
                                cellContent = String.valueOf(cell.getBooleanCellValue());
                                break;

                            case FORMULA:
                                cellContent = cell.getCellFormula();
                                break;

                            default:
                                break;
                        }

                        if (cellContent.equals(word)) {
                            int rowIndex = row.getRowNum() + 1;  // Excel row number (1-based)
                            int columnIndex = cell.getColumnIndex();  // Column index (0-based)

                            // Display sheet number, row number, and column
                            System.out.println("Found content in:");
                            System.out.println("Sheet: " + (i + 1));  // Sheet index (1-based)
                            System.out.println("Row: " + rowIndex);
                            System.out.println("--------------------------------");

                        }
                    }
                }


            }


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}