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
        String excelFilePath = "D:\\OneDrive - Olam International\\Desktop\\Amend-Quarantine\\Exact_Necessary_DATA\\PROD_DATA\\ExcelSheet-for-Calculated-Questions.xlsx";
        
        //D:\OneDrive - Olam International\Desktop\Amend-Quarantine\Exact_Necessary_DATA\PROD_DATA\ExcelSheet-for-Non-Calculated-Questions.xlsx
        //D:\OneDrive - Olam International\Desktop\Amend-Quarantine\Exact_Necessary_DATA\PROD_DATA\ExcelSheet-for-Calculated-Questions.xlsx
  
        
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            Iterator<Row> iterator = sheet.iterator();

            // Skip the header row
            if (iterator.hasNext()) {
                iterator.next();
            }

            List<Integer> submittedModuleIdList = new ArrayList<>();
            List<Integer> answerIdList = new ArrayList<>();
            List<Integer> updatedAnswersList = new ArrayList<>();
            List<Integer> questionIdList = new ArrayList<>();

            int jsonDataCount = 0;
            JSONArray jsonArray = new JSONArray();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                submittedModuleIdList.add((int) currentRow.getCell(1).getNumericCellValue());
                answerIdList.add((int) currentRow.getCell(2).getNumericCellValue());
                updatedAnswersList.add((int) currentRow.getCell(3).getNumericCellValue());
                questionIdList.add((int) currentRow.getCell(0).getNumericCellValue());

                jsonDataCount++;

                if (jsonDataCount == 100) {
                    // Create a JSON object and reset the lists for the next 100 entries
                    JSONObject jsonObject = new JSONObject();
                    jsonObject.put("submittedModuleIdList", new JSONArray(submittedModuleIdList));
                    jsonObject.put("answerIdList", new JSONArray(answerIdList));
                    jsonObject.put("updatedAnswersList", new JSONArray(updatedAnswersList));
                    jsonObject.put("questionIdList", new JSONArray(questionIdList));

                    jsonArray.put(jsonObject);

                    // Reset lists and counter
                    submittedModuleIdList.clear();
                    answerIdList.clear();
                    updatedAnswersList.clear();
                    questionIdList.clear();
                    jsonDataCount = 0;
                }
            }

            // If there are remaining entries (less than 100), create the last JSON object
            if (!submittedModuleIdList.isEmpty()) {
                JSONObject jsonObject = new JSONObject();
                jsonObject.put("submittedModuleIdList", new JSONArray(submittedModuleIdList));
                jsonObject.put("answerIdList", new JSONArray(answerIdList));
                jsonObject.put("updatedAnswersList", new JSONArray(updatedAnswersList));
                jsonObject.put("questionIdList", new JSONArray(questionIdList));

                jsonArray.put(jsonObject);
            }

            System.out.println(jsonArray.toString(2));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

