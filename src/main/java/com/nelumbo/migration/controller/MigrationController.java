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

    @PostMapping("empresa")
    public ResponseEntity<Void> migrateEmpresa(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateEmpresa(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("region")
    public ResponseEntity<Void> migrateRegion(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateRegion(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("division")
    public ResponseEntity<Void> migrateDivision(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateDivision(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("zona")
    public ResponseEntity<Void> migrateZona(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateZona(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("area")
    public ResponseEntity<Void> migrateArea(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateArea(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("subarea")
    public ResponseEntity<Void> migrateSubarea(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateSubarea(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("departamento")
    public ResponseEntity<Void> migrateDepartamento(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateDepartamento(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("cost-centers")
    public ResponseEntity<Void> migrateCostCenters(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateCostCenters(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("cost-centers-org-entities-geographic")
    public ResponseEntity<Void> migrateCostCentersOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateCostCentersOrgEntitiesGeographic(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("cost-centers-org-entities-organizative")
    public ResponseEntity<Void> migrateCostCentersOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateCostCentersOrgEntitiesOrganizative(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("stores")
    public ResponseEntity<Void> migrateStores(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStores(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("stores-org-entities-geographic")
    public ResponseEntity<Void> migrateStoresOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStoresOrgEntitiesGeographic(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("stores-org-entities-organizative")
    public ResponseEntity<Void> migrateStoresOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateStoresOrgEntitiesOrganizative(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("work-positions")
    public ResponseEntity<Void> migrateWorkPositions(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateWorkPositions(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("work-positions-details")
    public ResponseEntity<Void> migrateWorkPositionsDetails(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateWorkPositionsDetails(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("profiles")
    public ResponseEntity<Void> migrateProfiles(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateProfiles(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("profiles-groups")
    public ResponseEntity<Void> migrateProfilesGroups(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateProfilesGroups(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("references")
    public ResponseEntity<Void> migrateReferences(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateReferences(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("info-bancaria")
    public ResponseEntity<Void> migrateInfoBancaria(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateInfoBancaria(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("info-sueldos")
    public ResponseEntity<Void> migrateInfoSueldos(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateInfoSueldos(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    @PostMapping("/load-compensations")
    public ResponseEntity<Void> cargarCompensaciones(@RequestPart(value = "file") MultipartFile file) {
        migrationService.loadCompensationsCategories(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("/load-tabs")
    public ResponseEntity<Void> loadTabs(@RequestPart(value = "file") MultipartFile file) {
        migrationService.loadTabs(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }

    @PostMapping("/load-work-position-categories")
    public ResponseEntity<Void> loadWorkPositionCategories(@RequestPart(value = "file") MultipartFile file) {
        migrationService.loadWorkPositionCategories(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
    
    @PostMapping("/load-groups")
    public ResponseEntity<Void> loadWorkGroups(@RequestPart(value = "file") MultipartFile file) {
        migrationService.loadGroups(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
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
