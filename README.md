# ExcelManipulator



import com.excelImporter.service.RTIMultilingualService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;


@RestController
@RequestMapping("/api/multilingual")
public class RTIMultilingualController {

    private final RTIMultilingualService service;

    public RTIMultilingualController(RTIMultilingualService service) {
        this.service = service;
    }

    private static final Logger log = LoggerFactory.getLogger(RTIMultilingualController.class);

    @PostMapping("/upload")
    public ResponseEntity<Void> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            service.uploadExcel(file);  // process file as usual
            return ResponseEntity.ok().build();  // 200 OK with empty body
        } catch (Exception e) {
            log.error("Error while processing Excel file: {}", file.getOriginalFilename(), e);
            return ResponseEntity.status(500).build(); // 500 if something went wrong
        }
    }
}


-------------------





import com.excelImporter.controller.RTIMultilingualController;
import com.excelImporter.dao.RTIMultilingualDao;
import com.excelImporter.domain.RTIMultilingual;
import com.excelImporter.dto.RTIMultilingualDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class RTIMultilingualService {

    private final RTIMultilingualDao repository;

    public RTIMultilingualService(RTIMultilingualDao repository) {
        this.repository = repository;
    }

    private static final Logger log = LoggerFactory.getLogger(RTIMultilingualService.class);

    public List<RTIMultilingualDTO> uploadExcel(MultipartFile file) throws Exception {
        List<RTIMultilingualDTO> savedList = new ArrayList<>();

        log.info("Starting Excel upload: {}", file.getOriginalFilename());

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // skip header
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Long id = null;
                if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC) {
                    id = (long) row.getCell(0).getNumericCellValue();
                }

                Long reqformId = (long) row.getCell(1).getNumericCellValue();
                String label = row.getCell(2).getStringCellValue();

                // langId is string now
                String langId = row.getCell(3).getStringCellValue();

                RTIMultilingual entity = (id != null)
                        ? repository.findById(id).orElse(new RTIMultilingual())
                        : new RTIMultilingual();

                entity.setReqformId(reqformId);
                entity.setLabel(label);
                entity.setLangId(langId);

                RTIMultilingual savedEntity = repository.save(entity);

                // Log message
                if (id != null) {
                    log.info("Updated record with id={}", savedEntity.getId());
                } else {
                    log.info("Created new record with id={}", savedEntity.getId());
                }
                RTIMultilingualDTO dto = new RTIMultilingualDTO();
                dto.setId(savedEntity.getId());
                dto.setReqformId(savedEntity.getReqformId());
                dto.setLabel(savedEntity.getLabel());
                dto.setLangId(savedEntity.getLangId());

                savedList.add(dto);
            }

            workbook.close();
        }
        log.info("Excel upload completed. Total records processed: {}", savedList.size());

        return savedList;
    }
}
------------------


public interface RTIMultilingualDao extends JpaRepository<RTIMultilingual, Long> {

}
-------------



public class RTIMultilingualDTO {

    private Long id;        // nullable for create
    private Long reqformId;
    private String label;
    private String langId;

    public RTIMultilingualDTO() {
    }
    public RTIMultilingualDTO(Long id, Long reqformId, String label, String langId) {
        this.id = id;
        this.reqformId = reqformId;
        this.label = label;
        this.langId = langId;
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
}
------------------------


@Table(name ="rti_multilingual")
@Entity
public class RTIMultilingual {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;        // nullable for create
    private Long reqformId;
    private String label;
    private String langId;

    public RTIMultilingual() {
    }
    public RTIMultilingual(Long id, Long reqformId, String label, String langId) {
        this.id = id;
        this.reqformId = reqformId;
        this.label = label;
        this.langId = langId;
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
}
