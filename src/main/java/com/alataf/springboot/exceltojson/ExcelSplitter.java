package com.alataf.springboot.exceltojson;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelSplitter {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\alataf.ansari\\Downloads\\GHANA UR 3.8.23-updated.xlsx";
        String outputFilePath1 = "D:\\OneDrive - Olam International\\Desktop\\Amend-Quarantine\\ExcelSheet1.xlsx";
        String outputFilePath2 = "D:\\OneDrive - Olam International\\Desktop\\Amend-Quarantine\\ExcelSheet2.xlsx";
        
        //D:\OneDrive - Olam International\Desktop\Amend-Quarantine\ExcelSheet1.xlsx
        //D:\OneDrive - Olam International\Desktop\Amend-Quarantine\ExcelSheet2.xlsx
  

        List<Integer> sheet1QuestionIds = Arrays.asList(16, 50);
        List<Integer> sheet2QuestionIds = Arrays.asList(7, 20, 3854, 3855, 3856, 3857);

        try {
            splitExcelSheet(inputFilePath, outputFilePath1, outputFilePath2, sheet1QuestionIds, sheet2QuestionIds);
            System.out.println("Excel sheets created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void splitExcelSheet(String inputFilePath, String outputFilePath1, String outputFilePath2,
                                        List<Integer> sheet1QuestionIds, List<Integer> sheet2QuestionIds) throws IOException {
        FileInputStream inputStream = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);

        Sheet inputSheet = workbook.getSheetAt(1);

        Workbook outputWorkbook1 = new XSSFWorkbook();
        Workbook outputWorkbook2 = new XSSFWorkbook();

        Sheet outputSheet1 = outputWorkbook1.createSheet("Sheet1");
        Sheet outputSheet2 = outputWorkbook2.createSheet("Sheet2");

        // Create headers in the output sheets
        Row headerRow = outputSheet1.createRow(0);
        copyRow(inputSheet.getRow(0), headerRow);
        Row headerRow2 = outputSheet2.createRow(0);
        copyRow(inputSheet.getRow(0), headerRow2);

        int rowIndex1 = 1; // Start from the second row in the output sheets
        int rowIndex2 = 1;

        for (int rowIndex = 1; rowIndex <= inputSheet.getLastRowNum(); rowIndex++) {
            Row currentRow = inputSheet.getRow(rowIndex);
            if (currentRow != null) {
                Cell questionIdCell = currentRow.getCell(0);
                if (questionIdCell != null && questionIdCell.getCellType() == CellType.NUMERIC) {
                    int questionId = (int) questionIdCell.getNumericCellValue();

                    if (sheet1QuestionIds.contains(questionId)) {
                        copyRow(currentRow, outputSheet1.createRow(rowIndex1++));
                    } else if (sheet2QuestionIds.contains(questionId)) {
                        copyRow(currentRow, outputSheet2.createRow(rowIndex2++));
                    }
                }
            }
        }

        // Write the output to files
        try (FileOutputStream outputStream1 = new FileOutputStream(outputFilePath1);
             FileOutputStream outputStream2 = new FileOutputStream(outputFilePath2)) {
            outputWorkbook1.write(outputStream1);
            outputWorkbook2.write(outputStream2);
        }

        // Close all resources
        inputStream.close();
        outputWorkbook1.close();
        outputWorkbook2.close();
        workbook.close();
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        if (sourceRow == null || targetRow == null) {
            return;
        }

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.createCell(i, sourceCell == null ? CellType.BLANK : sourceCell.getCellType());

            if (sourceCell != null) {
                if (sourceCell.getCellType() == CellType.NUMERIC) {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                } else if (sourceCell.getCellType() == CellType.STRING) {
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                } else if (sourceCell.getCellType() == CellType.BOOLEAN) {
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                }
                // Add conditions for other cell types if necessary
            }
        }
    }
}

