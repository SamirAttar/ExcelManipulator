
import com.XlsImporter.XlsImporter.service.RTiReqFormFieldMultilingualService;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiImplicitParams;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping(RTiReqFormFieldMultilingualController.BASE_PATH)
public class RTiReqFormFieldMultilingualController {

    @Autowired
    private RTiReqFormFieldMultilingualService service;

    public static final String BASE_PATH = "api/v1/rti";

    @PostMapping(value = "/{categoryType}")
    @ApiOperation(
            value = "Upload Excel File for RTI Multilingual Data",
            nickname = "uploadRTIMultilingualExcel",
            notes = "Creates new records if they do not exist and updates only the label if reqFormFieldID and languageID already exist."
    )
    @ApiImplicitParams({
            @ApiImplicitParam(name = "orgoid", value = "Organization ID", dataType = "string", paramType = "header", required = true),
            @ApiImplicitParam(name = "associateoid", value = "Associate ID", dataType = "string", paramType = "header", required = true)
    })
    public ResponseEntity<String> uploadExcel(
            @RequestHeader("orgoid") String orgOid,
            @RequestHeader("associateoid") String associateOid,
            @RequestParam("file") MultipartFile file,
            @PathVariable String categoryType
    ) {
        try {
            service.uploadExcel(file);
            return ResponseEntity.ok().build();  // 200 OK with no body
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build(); // 500 Internal Server Error
        }
    }
}



-------------------




package com.XlsImporter.XlsImporter.service;

import com.XlsImporter.XlsImporter.domian.RTiReqFormFieldMultilingual;
import com.XlsImporter.XlsImporter.repository.RTiReqFormFieldMultilingualRepository;
import jakarta.transaction.Transactional;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.*;

@Service
public class RTiReqFormFieldMultilingualService {

    @Autowired
    private RTiReqFormFieldMultilingualRepository rTIReqFormFieldMultilingualRepository;

    @Transactional
    public void uploadExcel(MultipartFile file) throws Exception {
        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            // Read header row to get languageID columns (skip ID & Label)
            Row headerRow = sheet.getRow(0);
            Map<Integer, String> languageMap = new HashMap<>();
            for (int i = 2; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && !cell.getStringCellValue().trim().isEmpty()) {
                    languageMap.put(i, cell.getStringCellValue());
                }
            }

            // Iterate through data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell idCell = row.getCell(0);
                if (idCell == null) continue;
                Long reqFormFieldID = (long) idCell.getNumericCellValue();

                for (Map.Entry<Integer, String> entry : languageMap.entrySet()) {
                    int colIndex = entry.getKey();
                    String languageID = entry.getValue();

                    Cell labelCell = row.getCell(colIndex);
                    if (labelCell == null) continue;

                    String label = labelCell.getStringCellValue();
                    if (label == null || label.trim().isEmpty()) continue; // skip empty labels

                    Optional<RTiReqFormFieldMultilingual> existing = rTIReqFormFieldMultilingualRepository
                            .findByReqFormFieldIDAndLanguageID(reqFormFieldID, languageID);

                    if (existing.isPresent()) {
                        RTiReqFormFieldMultilingual entity = existing.get();
                        entity.setLabel(label); // update only label
                        rTIReqFormFieldMultilingualRepository.save(entity);
                    } else {
                        RTiReqFormFieldMultilingual entity = new RTiReqFormFieldMultilingual();
                        entity.setReqFormFieldID(reqFormFieldID);
                        entity.setLanguageID(languageID);
                        entity.setLabel(label);
                        rTIReqFormFieldMultilingualRepository.save(entity);
                    }
                }
            }
        }
    }

}

---------

public interface RTiReqFormFieldMultilingualRepository extends JpaRepository<RTiReqFormFieldMultilingual, Long>{

    Optional<RTiReqFormFieldMultilingual> findByReqFormFieldIDAndLanguageID(Long reqFormFieldID, String languageID);


}
