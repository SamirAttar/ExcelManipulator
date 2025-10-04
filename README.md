# ExcelManipulator



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

