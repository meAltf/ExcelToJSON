package com.alataf.springboot.exceltojson;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class JsonConverter {

    public static void main(String[] args) {

        //The exact path of the file
        String excelFilePath = "exact/path/inputfile.xlsx";
        
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            Iterator<Row> iterator = sheet.iterator();

            // Skip the header row
            if (iterator.hasNext()) {
                iterator.next();
            }

            List<Integer> columnName1 = new ArrayList<>();
            List<Integer> columnName2 = new ArrayList<>();
            List<Integer> columnName3 = new ArrayList<>();
            List<Integer> columnName4 = new ArrayList<>();

            int jsonDataCount = 0;
            JSONArray jsonArray = new JSONArray();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                columnName1.add((int) currentRow.getCell(1).getNumericCellValue());
                columnName2.add((int) currentRow.getCell(2).getNumericCellValue());
                columnName3.add((int) currentRow.getCell(3).getNumericCellValue());
                columnName4.add((int) currentRow.getCell(0).getNumericCellValue());

                jsonDataCount++;

                if (jsonDataCount == 100) {
                    // Create a JSON object and reset the lists for the next 100 entries
                    JSONObject jsonObject = new JSONObject();
                    jsonObject.put("columnName1", new JSONArray(columnName1));
                    jsonObject.put("columnName2", new JSONArray(columnName2));
                    jsonObject.put("columnName3", new JSONArray(columnName3));
                    jsonObject.put("columnName4", new JSONArray(columnName4));

                    jsonArray.put(jsonObject);

                    // Reset lists and counter
                    columnName1.clear();
                    columnName2.clear();
                    columnName3.clear();
                    columnName4.clear();
                    jsonDataCount = 0;
                }
            }

            // If there are remaining entries (less than 100), create the last JSON object
            if (!columnName1.isEmpty()) {
                JSONObject jsonObject = new JSONObject();
                jsonObject.put("columnName1", new JSONArray(columnName1));
                jsonObject.put("columnName2", new JSONArray(columnName2));
                jsonObject.put("columnName3", new JSONArray(columnName3));
                jsonObject.put("columnName4", new JSONArray(columnName4));

                jsonArray.put(jsonObject);
            }

            System.out.println(jsonArray.toString(2));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

