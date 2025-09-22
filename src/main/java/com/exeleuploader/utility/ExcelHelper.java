package com.exeleuploader.utility;

import com.exeleuploader.Model.User;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelHelper {

    // ✅ Check if file is Excel
    public static boolean hasExcelFormat(MultipartFile file) {
        return file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    // ✅ Convert Excel rows into Users
    public static List<User> excelToUsers(MultipartFile file) {
        try {
            List<User> users = new ArrayList<>();
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rows = sheet.iterator();
            int rowNumber = 0;

            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // Skip header row
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                User user = new User();
                user.setUsername(currentRow.getCell(0).getStringCellValue());
                user.setCity(currentRow.getCell(1).getStringCellValue());
                user.setEmail(currentRow.getCell(2).getStringCellValue());

                users.add(user);
            }

            workbook.close();
            return users;

        } catch (IOException e) {
            throw new RuntimeException("Failed to parse Excel file: " + e.getMessage());
        }
    }
}
