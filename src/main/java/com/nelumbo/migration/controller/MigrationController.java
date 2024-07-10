package com.nelumbo.migration.controller;

import com.nelumbo.migration.service.MigrationService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

@Slf4j
@RestController
@RequestMapping("/migration")
@RequiredArgsConstructor
public class MigrationController {

    private static final String APPLICATION_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String ATTACHMENT_FILENAME = "attachment; filename=";
    private static final String ERROR_EXCEL = "Error returning modified Excel file: ";

    private final MigrationService migrationService;

    @PostMapping("cost-centers")
    public ResponseEntity<Void> migrateCostCenters(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateCostCenters(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("stores")
    public ResponseEntity<Void> migrateStores(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStores(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("stores-org-entities-details")
    public ResponseEntity<Void> migrateStoresOrgEntities(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStoresOrgEntities(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("stores-work-periods")
    public ResponseEntity<Void> migrateStoreWorkPeriods(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStoreWorkPeriods(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("work-positions")
    public ResponseEntity<Void> migrateWorkPositions(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateWorkPositions(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("profiles")
    public ResponseEntity<Void> migrateProfiles(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateProfiles(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("references")
    public ResponseEntity<Void> migrateReferences(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateReferences(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("/load-compensations")
    public ResponseEntity<InputStreamResource> cargarCompensaciones(@RequestPart(value = "file") MultipartFile file) {
        File modifiedFile = migrationService.loadCompensationsCategories(file);
        return  processFile(modifiedFile);
    }

    @PostMapping("/load-tabs")
    public ResponseEntity<InputStreamResource> loadTabs(@RequestPart(value = "file") MultipartFile file) {
        File modifiedFile = migrationService.loadTabs(file);
        return  processFile(modifiedFile);
    }

    @PostMapping("/load-work-position-categories")
    public ResponseEntity<InputStreamResource> loadWorkPositionCategories(@RequestPart(value = "file") MultipartFile file) {
        File modifiedFile = migrationService.loadWorkPositionCategories(file);
        return  processFile(modifiedFile);
    }

    @PostMapping("/load-work-periods")
    public ResponseEntity<InputStreamResource> loadWorkPeriods(@RequestPart(value = "file") MultipartFile file) {
        File modifiedFile = migrationService.loadWorkPeriods(file);
        return  processFile(modifiedFile);
    }

    private ResponseEntity<InputStreamResource> processFile(File modifiedFile) {
        try {
            InputStreamResource resource = new InputStreamResource(new FileInputStream(modifiedFile));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, ATTACHMENT_FILENAME + modifiedFile.getName());

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(modifiedFile.length())
                    .contentType(MediaType.parseMediaType(APPLICATION_EXCEL))
                    .body(resource);
        } catch (IOException e) {
            log.error(ERROR_EXCEL + " {}", e.getMessage());
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
