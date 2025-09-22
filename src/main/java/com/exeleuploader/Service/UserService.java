package com.exeleuploader.Service;

import com.exeleuploader.Model.User;
import com.exeleuploader.Repo.UserRepo;
import com.exeleuploader.utility.ExcelHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

@Service
public class UserService {

    @Autowired
    private UserRepo userRepository;

    private static final Logger logger = LoggerFactory.getLogger(UserService.class);


    public String saveUpdateFromExcel(MultipartFile file) throws Exception {
        int updated = 0, created = 0;

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            logger.info("Processing Excel file: {}", file.getOriginalFilename());

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Read ID
                Long id = null;
                if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC) {
                    id = (long) row.getCell(0).getNumericCellValue();
                }

                // Read other values
                String name = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : "";
                String city = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "";
                String email = row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "";

                User user = null;

                if (id != null && userRepository.existsById(id)) {
                    // Case 1: Update by ID
                    user = userRepository.findById(id).get();
                    logger.info("Updating user with ID {}", id);
                    updated++;
                } else if (!email.isBlank()) {
                    // Case 2: Check by email if ID not given
                    user = userRepository.findByEmail(email).orElse(null);
                    if (user != null) {
                        logger.info("Updating user by email: {}", email);
                        updated++;
                    }
                }

                if (user == null) {
                    // Case 3: Create new user
                    user = new User();
                    logger.info("Creating new user: {}", name);
                    created++;
                }

                // Update values
                user.setUsername(name);
                user.setCity(city);
                user.setEmail(email);

                userRepository.save(user);
            }
        }

        return "Upload completed: " + updated + " updated, " + created + " created.";
    }














    @Service
    public class UserService {

        @Autowired
        private UserRepository userRepository;

        /**
         * Reads an Excel file and saves/updates users in the database.
         * @param file Excel file uploaded by the user
         * @return success message
         * @throws Exception if there is an error reading the file
         */
        public String saveUpdate2FromExcel(MultipartFile file) throws Exception {
            // Open the Excel file
            try (InputStream inputStream = file.getInputStream();
                 Workbook workbook = new XSSFWorkbook(inputStream)) {

                // Get the first sheet
                Sheet sheet = workbook.getSheetAt(0);

                // Loop through all rows, starting from row 1 (skip header)
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue; // skip empty rows

                    // Read values from cells
                    String name = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : "";
                    String city = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "";
                    String email = row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "";

                    if (email.isBlank()) continue; // skip rows with no email

                    // Check if user already exists by email
                    // If exists, fetch user; else create new user
                    User user = userRepository.findByEmail(email).orElse(new User());

                    // Set/update user fields
                    user.setName(name);
                    user.setCity(city);
                    user.setEmail(email);

                    // Save user to database
                    userRepository.save(user);
                }
            }

            return "Upload completed successfully!";
        }























    public void saveUsersFromExcel(MultipartFile file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // skip header row
                Row row = sheet.getRow(i);
                if (row == null) continue;

                User user = new User();
                user.setUsername(row.getCell(0).getStringCellValue());
                user.setCity(row.getCell(1).getStringCellValue());
                user.setEmail(row.getCell(2).getStringCellValue());

                userRepository.save(user);
            }
        }
    }

    public byte[] exportUsersToExcel() {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Users");

            // Header row
            Row headerRow = sheet.createRow(0);
            String[] headers = {"ID", "Username", "City", "Email"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Data rows
            List<User> users = userRepository.findAll();
            int rowIdx = 1;
            for (User user : users) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUsername());
                row.createCell(2).setCellValue(user.getCity());
                row.createCell(3).setCellValue(user.getEmail());
            }

            workbook.write(out);
            return out.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("Error generating Excel file: " + e.getMessage(), e);
        }
    }
}