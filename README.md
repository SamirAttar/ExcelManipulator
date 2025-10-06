# ExcelManipulator



package com.excelImporter.service;

import com.excelImporter.dao.RTiReqFormFieldMultilingualRepository;
import com.excelImporter.domain.RTiReqFormFieldMultilingual;
import com.excelImporter.utils.SiteTranslator;
import jakarta.transaction.Transactional;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;

@Slf4j
@Service
public class RTiReqFormFieldMultilingualService {

    @Autowired
    private RTiReqFormFieldMultilingualRepository repository;

    @Autowired
    private SiteTranslator siteTranslator;


    private static final Logger log = LoggerFactory.getLogger(RTiReqFormFieldMultilingualService.class);

    @Transactional
    public void uploadExcel(MultipartFile file) throws Exception {
        log.info("Starting Excel upload: {}", file.getOriginalFilename());

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            if (sheet.getPhysicalNumberOfRows() < 2) {
                log.warn("Excel file has no data rows.");
                throw new RuntimeException("Excel file has no data rows.");
            }

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                log.warn("Excel file has no header row.");
                throw new RuntimeException("Excel file has no header row.");
            }

            int idColumnIndex = -1;
            int labelColumnIndex = -1;
            Map<Integer, String> languageColumns = new HashMap<>();

            // --- Detect column indexes ---
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell == null) continue;
                cell.setCellType(CellType.STRING);
                String header = cell.getStringCellValue().trim().toUpperCase();

                switch (header) {
                    case "ID" -> {
                        idColumnIndex = i;
                        log.info("Found ID column at index {}", i);
                    }
                    case "LABLE" -> {
                        labelColumnIndex = i;
                        log.info("Found LABLE column at index {}", i);
                    }
                    default -> {
                        String languageID = siteTranslator.getLanguageIdFromName(header);
                        if (languageID != null && !languageID.isEmpty()) {
                            languageColumns.put(i, languageID);
                            log.info("Mapped Excel header '{}' to languageID '{}'", header, languageID);
                        } else {
                            log.warn("Skipping unknown language column '{}'", header);
                        }
                    }
                }
            }

            if (idColumnIndex == -1) {
                log.error("Missing ID column in Excel");
                throw new RuntimeException("Missing ID column.");
            }
            if (languageColumns.isEmpty()) {
                log.error("No editable language columns found in Excel");
                throw new RuntimeException("No editable language columns found.");
            }

            // --- Iterate data rows ---
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    log.info("Skipping empty row {}", i);
                    continue;
                }

                // --- Read reqFormFieldID ---
                Cell idCell = row.getCell(idColumnIndex);
                if (idCell == null) {
                    log.warn("Skipping row {}: ID cell is null", i);
                    continue;
                }

                Long reqFormFieldID = null;
                if (idCell.getCellType() == CellType.NUMERIC) {
                    reqFormFieldID = (long) idCell.getNumericCellValue();
                } else {
                    try {
                        reqFormFieldID = Long.parseLong(idCell.getStringCellValue().trim());
                    } catch (Exception ignored) {}
                }
                if (reqFormFieldID == null) {
                    log.warn("Skipping row {}: Invalid ID value", i);
                    continue;
                }

                // --- Process all editable language columns ---
                for (Map.Entry<Integer, String> entry : languageColumns.entrySet()) {
                    int colIndex = entry.getKey();
                    String languageID = entry.getValue();

                    Cell langCell = row.getCell(colIndex);
                    if (langCell == null) continue;

                    langCell.setCellType(CellType.STRING);
                    String label = langCell.getStringCellValue().trim();
                    if (label.isEmpty()) continue;

                    Optional<RTiReqFormFieldMultilingual> existingOpt =
                            repository.findByReqFormFieldIDAndLanguageID(reqFormFieldID, languageID);

                    RTiReqFormFieldMultilingual entity;
                    if (existingOpt.isPresent()) {
                        entity = existingOpt.get();
                        log.info("Updating label for reqFormFieldID={} and languageID={}", reqFormFieldID, languageID);
                    } else {
                        entity = new RTiReqFormFieldMultilingual();
                        entity.setReqFormFieldID(reqFormFieldID);
                        entity.setLanguageID(languageID);
                        log.info("Creating new entity for reqFormFieldID={} and languageID={}", reqFormFieldID, languageID);
                    }

                    entity.setLabel(label); // update label
                    repository.save(entity);
                }
            }

            workbook.close();
            log.info("Excel upload completed successfully: {}", file.getOriginalFilename());
        }
    }
}


-------

import com.excelImporter.domain.RTiReqFormFieldMultilingual;
import org.springframework.data.jpa.repository.JpaRepository;



import java.util.Optional;

public interface RTiReqFormFieldMultilingualRepository extends JpaRepository<RTiReqFormFieldMultilingual, Long> {

    Optional<RTiReqFormFieldMultilingual> findByReqFormFieldIDAndLanguageID(Long reqFormFieldID, String languageID);
}



////////



















public interface RTiReqFormFieldMultilingualRepository extends JpaRepository<RTiReqFormFieldMultilingual, Long> {
    Optional<RTiReqFormFieldMultilingual> findByReqFormFieldIdAndLanguageId(Long reqFormFieldId, String languageId);
}

package com.excelImporter.service;

import com.adp.nas.tas.rm.translations.domain.RTiReqFormFieldMultilingual;
import com.adp.nas.tas.rm.translations.repositories.RTiReqFormFieldMultilingualRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@Slf4j
@Service
@RequiredArgsConstructor
public class RTIMultilingualService {

    private final RTiReqFormFieldMultilingualRepository repository;

    public void uploadExcel(MultipartFile file) {
        log.info("Starting Excel upload: {}", file.getOriginalFilename());

        try (InputStream is = file.getInputStream();
             Workbook workbook = WorkbookFactory.create(is)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Skip header row
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Long reqFormFieldId = getLongValue(row.getCell(1));
                String label = getStringValue(row.getCell(2));
                String languageId = getStringValue(row.getCell(3));
                String help = getStringValue(row.getCell(4));

                // Skip invalid rows
                if (reqFormFieldId == null || languageId.isEmpty()) {
                    log.warn("Skipping invalid row {} - missing reqFormFieldId or languageId", i);
                    continue;
                }

                repository.findByReqFormFieldIdAndLanguageId(reqFormFieldId, languageId)
                        .ifPresentOrElse(existing -> {
                            // Update only if label and help are not empty
                            if (!label.isEmpty() && !help.isEmpty()) {
                                existing.setLabel(label);
                                existing.setHelp(help);
                                repository.save(existing);
                                log.info("Updated existing record: reqFormFieldId={}, languageId={}", reqFormFieldId, languageId);
                            } else {
                                log.warn("Skipped update for row {} - empty label/help", i);
                            }
                        }, () -> {
                            // Insert new record
                            RTiReqFormFieldMultilingual newRecord = new RTiReqFormFieldMultilingual();
                            newRecord.setReqFormFieldId(reqFormFieldId);
                            newRecord.setLabel(label);
                            newRecord.setLanguageId(languageId);
                            newRecord.setHelp(help);
                            repository.save(newRecord);
                            log.info("Inserted new record: reqFormFieldId={}, languageId={}", reqFormFieldId, languageId);
                        });
            }

            log.info("✅ Excel file processed successfully: {}", file.getOriginalFilename());

        } catch (Exception e) {
            log.error("❌ Error while processing Excel file: {}", file.getOriginalFilename(), e);
            throw new RuntimeException("Failed to process Excel file", e);
        }
    }

    // Helper methods to handle Excel cells safely
    private String getStringValue(Cell cell) {
        if (cell == null) return "";
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue().trim();
    }

    private Long getLongValue(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return (long) cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty()) {
                return Long.parseLong(cell.getStringCellValue().trim());
            }
        } catch (Exception e) {
            log.warn("Invalid numeric value in cell: {}", cell, e);
        }
        return null;
    }
}

new changes 



package com.excelImporter.controller;

import com.excelImporter.service.RTIMultilingualService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/api/v1/rti")
public class RTIMultilingualController {

    private final RTIMultilingualService rtiMultilingualService;

    @PostMapping("/upload")
    public ResponseEntity<String> uploadExcel(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "orgoid", required = false) String orgOid,
            @RequestHeader(value = "associateoid", required = false) String associateOid) {

        log.info("Received Excel upload request from org: {}, associate: {}", orgOid, associateOid);

        if (file.isEmpty()) {
            log.warn("Uploaded file is empty");
            return ResponseEntity.badRequest().body("File is empty");
        }

        try {
            rtiMultilingualService.uploadExcel(file);
            return ResponseEntity.ok("Excel uploaded and processed successfully");
        } catch (Exception e) {
            log.error("Error while uploading Excel: {}", e.getMessage(), e);
            return ResponseEntity.internalServerError()
                    .body("Failed to process Excel file: " + e.getMessage());
        }
    }
}

---------

package com.excelImporter.service;

import com.excelImporter.entity.RTIMultilingual;
import com.excelImporter.repository.RTIMultilingualRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@Slf4j
@Service
@RequiredArgsConstructor
public class RTIMultilingualService {

    private final RTIMultilingualRepository repository;

    public void uploadExcel(MultipartFile file) {
        log.info("Starting Excel upload: {}", file.getOriginalFilename());

        try (InputStream is = file.getInputStream(); Workbook workbook = WorkbookFactory.create(is)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Skip header (start from 1)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Long reqformId = getLongValue(row.getCell(1));
                String label = getStringValue(row.getCell(2));
                String langId = getStringValue(row.getCell(3));
                String help = getStringValue(row.getCell(4));

                if (reqformId == null || langId.isEmpty()) {
                    log.warn("Skipping invalid row {} - missing reqformId/langId", i);
                    continue;
                }

                // ✅ Check if record exists (reqformId + langId)
                repository.findByReqformIdAndLangId(reqformId, langId)
                        .ifPresentOrElse(existing -> {
                            // Update label only
                            existing.setLabel(label);
                            repository.save(existing);
                            log.info("Updated existing record: reqformId={}, langId={}", reqformId, langId);
                        }, () -> {
                            // Insert new record
                            RTIMultilingual newRecord = new RTIMultilingual();
                            newRecord.setReqformId(reqformId);
                            newRecord.setLabel(label);
                            newRecord.setLangId(langId);
                            newRecord.setHelp(help);
                            repository.save(newRecord);
                            log.info("Inserted new record: reqformId={}, langId={}", reqformId, langId);
                        });
            }

            log.info("✅ Excel file processed successfully: {}", file.getOriginalFilename());

        } catch (Exception e) {
            log.error("❌ Error while processing Excel file: {}", file.getOriginalFilename(), e);
            throw new RuntimeException("Failed to process Excel file", e);
        }
    }

    // Helper methods to handle different Excel cell types safely
    private String getStringValue(Cell cell) {
        if (cell == null) return "";
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue().trim();
    }

    private Long getLongValue(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return (long) cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty()) {
                return Long.parseLong(cell.getStringCellValue().trim());
            }
        } catch (Exception e) {
            log.warn("Invalid numeric value in cell: {}", cell, e);
        }
        return null;
    }
}



--------

@Slf4j
@RestController
@RequestMapping(RTIMultilingualController.BASE_PATH)

public class RTIMultilingualController {

    public static final String BASE_PATH = "api/v1/rti";

    @Autowired
    private RTIMultilingualService service;

    @PostMapping("/upload")
    @ApiOperation(
            value = "Upload Excel File for RTI Multilingual Data",
            nickname = "uploadRTIMultilingualExcel",
            notes = "This API uploads an Excel file to create or update RTI Multilingual records in the database."
    )
    @ApiImplicitParams({
            @ApiImplicitParam(name = "orgoid", value = "Organization ID", dataType = "string", paramType = "header", required = true),
            @ApiImplicitParam(name = "associateoid", value = "Associate ID", dataType = "string", paramType = "header", required = true)
    })
    public ResponseEntity<Void> uploadExcel(
            @RequestHeader("orgoid") String orgOid,
            @RequestHeader("associateoid") String associateOid,
            @RequestParam("file") MultipartFile file
    ) {
//        log.info("Received Excel upload request - OrgID: {}, AssociateID: {}, File: {}",
//                orgOid, associateOid, file.getOriginalFilename());

        try {
            service.uploadExcel(file);
//            log.info("File '{}' uploaded and processed successfully.", file.getOriginalFilename());
            return ResponseEntity.ok().build(); // ✅ returns 200 OK
        } catch (Exception e) {
//            log.error("Error processing Excel file '{}': {}", file.getOriginalFilename(), e.getMessage(), e);
            return ResponseEntity.internalServerError().build(); // ✅ returns 500
        }
    }
}

-----------------

package com.excelImporter.controller;


import com.excelImporter.service.RTIMultilingualService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import lombok.extern.slf4j.Slf4j;


@RestController
@RequestMapping("/api/multilingual")
public class RTIMultilingualController {

    private static final Logger log = LoggerFactory.getLogger(RTIMultilingualController.class);


    private final RTIMultilingualService service;

    public RTIMultilingualController(RTIMultilingualService service) {
        this.service = service;
    }

    @PostMapping("/upload")
    public ResponseEntity<Void> uploadExcel(@RequestParam("file") MultipartFile file) {
        log.info("Received Excel upload request: {}", file.getOriginalFilename());
        try {
            service.uploadExcel(file);
            return ResponseEntity.ok().build(); // ✅ only 200 OK
        } catch (Exception e) {
            log.error("Error while processing Excel file: {}", file.getOriginalFilename(), e);
            return ResponseEntity.status(500).build(); // 500 if error
        }
    }
}
-------------------





package com.excelImporter.service;

import com.excelImporter.dao.RTIMultilingualDao;
import com.excelImporter.domain.RTIMultilingual;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;


@Service
public class RTIMultilingualService {

    private static final Logger log = LoggerFactory.getLogger(RTIMultilingualService.class);


    private final RTIMultilingualDao repository;

    public RTIMultilingualService(RTIMultilingualDao repository) {
        this.repository = repository;
    }

    public void uploadExcel(MultipartFile file) throws Exception {
        log.info("Starting Excel upload: {}", file.getOriginalFilename());

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // ID column
                Long id = null;
                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                    id = (long) idCell.getNumericCellValue();
                }

                // reqformId
                Long reqformId = null;
                Cell reqformCell = row.getCell(1);
                if (reqformCell != null) {
                    if (reqformCell.getCellType() == CellType.NUMERIC) {
                        reqformId = (long) reqformCell.getNumericCellValue();
                    } else {
                        try {
                            reqformId = Long.parseLong(reqformCell.getStringCellValue());
                        } catch (NumberFormatException ignored) {
                        }
                    }
                }

                // label
                String label = "";
                Cell labelCell = row.getCell(2);
                if (labelCell != null) {
                    labelCell.setCellType(CellType.STRING);
                    label = labelCell.getStringCellValue();
                }

                // langId
                String langId = "";
                Cell langCell = row.getCell(3);
                if (langCell != null) {
                    langCell.setCellType(CellType.STRING);
                    langId = langCell.getStringCellValue();
                }

                // help
                String help = "";
                Cell helpCell = row.getCell(4);
                if (helpCell != null) {
                    helpCell.setCellType(CellType.STRING);
                    help = helpCell.getStringCellValue();
                }

                // find existing or new entity
                RTIMultilingual entity = (id != null)
                        ? repository.findById(id).orElse(new RTIMultilingual())
                        : new RTIMultilingual();

                entity.setReqformId(reqformId);
                entity.setLabel(label);
                entity.setLangId(langId);
                entity.setHelp(help);

                RTIMultilingual savedEntity = repository.save(entity);

                log.info("{} record with id={} (row={})",
                        (id != null ? "Updated" : "Created"), savedEntity.getId(), i);
            }

            workbook.close();
        }

        log.info("Excel upload completed.");
    }
}

------------------

import com.excelImporter.domain.RTIMultilingual;
import org.springframework.data.jpa.repository.JpaRepository;


public interface RTIMultilingualDao extends JpaRepository<RTIMultilingual, Long> {

}

-------------



package com.excelImporter.dto;


import lombok.*;


public class RTIMultilingualDTO {

    private Long id;        // nullable for create
    private Long reqformId;
    private String label;
    private String langId;
    private String help;

    public RTIMultilingualDTO() {
    }
    public RTIMultilingualDTO(Long id, Long reqformId, String label, String langId, String help) {
        this.id = id;
        this.reqformId = reqformId;
        this.label = label;
        this.langId = langId;
        this.help = help;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Long getReqformId() {
        return reqformId;
    }

    public void setReqformId(Long reqformId) {
        this.reqformId = reqformId;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public String getLangId() {
        return langId;
    }

    public void setLangId(String langId) {
        this.langId = langId;
    }
    public String getHelp() {
        return help;
    }
    public void setHelp(String help) {
        this.help = help;
    }
}

------------------------



 package com.excelImporter.domain;


import jakarta.persistence.*;

@Table(name ="rti_multilingual")
@Entity
public class RTIMultilingual {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;        // nullable for create
    private Long reqformId;
    private String label;
    private String langId;
    private String help;

    public RTIMultilingual() {
    }

    public RTIMultilingual(Long id, Long reqformId, String label, String langId, String help) {
        this.id = id;
        this.reqformId = reqformId;
        this.label = label;
        this.langId = langId;
        this.help = help;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Long getReqformId() {
        return reqformId;
    }

    public void setReqformId(Long reqformId) {
        this.reqformId = reqformId;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public String getLangId() {
        return langId;
    }

    public void setLangId(String langId) {
        this.langId = langId;
    }
    public String getHelp() {
        return help;
    }
    public void setHelp(String help) {
        this.help = help;
    }
}


------
test

package com.excelImporter.controller;



import com.excelImporter.controller.RTIMultilingualController;
import com.excelImporter.service.RTIMultilingualService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.http.MediaType;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;

import java.io.ByteArrayOutputStream;

import static org.mockito.Mockito.doNothing;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@ExtendWith(MockitoExtension.class)
class RTIMultilingualControllerTest {

    private MockMvc mockMvc;

    @Mock
    private RTIMultilingualService service;

    @InjectMocks
    private RTIMultilingualController controller;

    @BeforeEach
    void setUp() {
        mockMvc = MockMvcBuilders.standaloneSetup(controller).build();
    }

    @Test
    void testUploadExcelEndpoint() throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("id");
        header.createCell(1).setCellValue("reqformId");
        header.createCell(2).setCellValue("label");
        header.createCell(3).setCellValue("langId");
        header.createCell(4).setCellValue("help");

        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue("");
        row.createCell(1).setCellValue(100);
        row.createCell(2).setCellValue("Label1");
        row.createCell(3).setCellValue("EN");
        row.createCell(4).setCellValue("Help text");

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();

        MockMultipartFile file = new MockMultipartFile(
                "file",
                "test.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                bos.toByteArray()
        );

        doNothing().when(service).uploadExcel(file);

        mockMvc.perform(multipart("/api/v1/rti/upload")
                        .file(file)
                        .header("orgoid", "ORG123")
                        .header("associateoid", "ASSO456")
                        .contentType(MediaType.MULTIPART_FORM_DATA))
                .andExpect(status().isOk());
    }
}
------

package com.excelImporter.service;

import com.excelImporter.dao.RTIMultilingualDao;
import com.excelImporter.domain.RTIMultilingual;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayOutputStream;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class RTIMultilingualServiceTest {

    @Mock
    private RTIMultilingualDao repository;

    @InjectMocks
    private RTIMultilingualService service;

    @Test
    void testUploadExcelWithInMemoryFile() throws Exception {
        // 1️⃣ Create in-memory Excel workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("id");
        header.createCell(1).setCellValue("reqformId");
        header.createCell(2).setCellValue("label");
        header.createCell(3).setCellValue("langId");
        header.createCell(4).setCellValue("help");

        // Data row
        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue(""); // id null → new record
        row.createCell(1).setCellValue(100);
        row.createCell(2).setCellValue("Label1");
        row.createCell(3).setCellValue("EN");
        row.createCell(4).setCellValue("Help text");

        // Write workbook to byte array
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();

        // 2️⃣ Wrap the byte array as MockMultipartFile
        MockMultipartFile file = new MockMultipartFile(
                "file",
                "test.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                bos.toByteArray()
        );

        // 3️⃣ Mock repository save behavior
        RTIMultilingual savedEntity = new RTIMultilingual();
        savedEntity.setId(1L);
        savedEntity.setReqformId(100L);
        savedEntity.setLabel("Label1");
        savedEntity.setLangId("EN");
        savedEntity.setHelp("Help text");

        when(repository.save(any(RTIMultilingual.class))).thenReturn(savedEntity);

        // 4️⃣ Call service method
        service.uploadExcel(file);

        // 5️⃣ Verify that repository.save() was called at least once
        verify(repository, atLeastOnce()).save(any(RTIMultilingual.class));
    }
}

--------


package com.excelImporter.dto;


import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class RTIMultilingualDtoTest {

    @Test
    void testDtoSettersAndGetters() {
        RTIMultilingualDTO dto = new RTIMultilingualDTO();
        dto.setId(1L);
        dto.setReqformId(100L);
        dto.setLabel("Label1");
        dto.setLangId("EN");
        dto.setHelp("Help text");

        assertThat(dto.getId()).isEqualTo(1L);
        assertThat(dto.getReqformId()).isEqualTo(100L);
        assertThat(dto.getLabel()).isEqualTo("Label1");
        assertThat(dto.getLangId()).isEqualTo("EN");
        assertThat(dto.getHelp()).isEqualTo("Help text");
    }
}


------


package com.excelImporter.dao;

import com.excelImporter.domain.RTIMultilingual;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.Optional;

import static org.mockito.Mockito.*;
import static org.assertj.core.api.Assertions.assertThat;

@ExtendWith(MockitoExtension.class)
class RTIMultilingualRepositoryTest {

    @Mock
    private RTIMultilingualDao repository;

    @Test
    void testSaveEntity() {
        RTIMultilingual entity = new RTIMultilingual();
        entity.setId(1L);
        entity.setReqformId(100L);
        entity.setLabel("Label1");
        entity.setLangId("EN");
        entity.setHelp("Help text");

        when(repository.save(entity)).thenReturn(entity);

        RTIMultilingual saved = repository.save(entity);

        assertThat(saved).isNotNull();
        assertThat(saved.getId()).isEqualTo(1L);
        assertThat(saved.getLabel()).isEqualTo("Label1");

        verify(repository, times(1)).save(entity);
    }

    @Test
    void testFindById() {
        RTIMultilingual entity = new RTIMultilingual();
        entity.setId(1L);
        entity.setLabel("Label1");

        when(repository.findById(1L)).thenReturn(Optional.of(entity));

        Optional<RTIMultilingual> found = repository.findById(1L);

        assertThat(found).isPresent();
        assertThat(found.get().getLabel()).isEqualTo("Label1");

        verify(repository, times(1)).findById(1L);
    }

    @Test
    void testExistsById() {
        when(repository.existsById(1L)).thenReturn(true);

        boolean exists = repository.existsById(1L);

        assertThat(exists).isTrue();

        verify(repository, times(1)).existsById(1L);
    }
}
------

package com.excelImporter.domain;

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class RTIMultilingualTest {

    @Test
    void testEntitySettersAndGetters() {
        RTIMultilingual entity = new RTIMultilingual();
        entity.setId(1L);
        entity.setReqformId(100L);
        entity.setLabel("Label1");
        entity.setLangId("EN");
        entity.setHelp("Help text");

        assertThat(entity.getId()).isEqualTo(1L);
        assertThat(entity.getReqformId()).isEqualTo(100L);
        assertThat(entity.getLabel()).isEqualTo("Label1");
        assertThat(entity.getLangId()).isEqualTo("EN");
        assertThat(entity.getHelp()).isEqualTo("Help text");
    }
}
