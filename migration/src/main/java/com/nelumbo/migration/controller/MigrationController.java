package com.nelumbo.migration.controller;

import com.nelumbo.migration.service.MigrationService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
@RestController
@RequestMapping("/migration")
@RequiredArgsConstructor
public class MigrationController {

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
}
