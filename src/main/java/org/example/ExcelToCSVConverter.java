package org.example;

import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelToCSVConverter {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\Hussien\\OneDrive\\Desktop\\x.xlsx";
        String csvFilePath = "C:\\Users\\Hussien\\OneDrive\\Desktop\\output.csv";

        convertExcelToCSV(excelFilePath, csvFilePath);
    }

    public static void convertExcelToCSV(String excelFilePath, String csvFilePath) {
        try {
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            FileWriter writer = new FileWriter(csvFilePath);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    CellType cellType = cell.getCellType();

                    if (cellType == CellType.STRING) {
                        writer.append(cell.getStringCellValue());
                    } else if (cellType == CellType.NUMERIC) {
                        writer.append(String.valueOf(cell.getNumericCellValue()));
                    } else if (cellType == CellType.BOOLEAN) {
                        writer.append(String.valueOf(cell.getBooleanCellValue()));
                    } else if (cellType == CellType.BLANK) {
                        writer.append(" - ");
                    } else {
                        // Handle other cell types as needed
                    }

                    writer.append(",");
                }
                writer.append("\n");
            }

            writer.flush();
            writer.close();
            workbook.close();
            inputStream.close();

            System.out.println("Excel file converted to CSV successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}