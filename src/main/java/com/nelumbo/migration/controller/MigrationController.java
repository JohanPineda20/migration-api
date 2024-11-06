package com.nelumbo.migration.controller;

import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;
import com.nelumbo.migration.service.MigrationService;
import jakarta.servlet.http.HttpServletRequest;
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
import java.util.List;

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
    public ResponseEntity<List<ErrorResponse>> migrateEmpresa(@RequestPart(value = "file") MultipartFile file,
                                               HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateEmpresa(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("region")
    public ResponseEntity<List<ErrorResponse>> migrateRegion(@RequestPart(value = "file") MultipartFile file,
                                                             HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateRegion(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("division")
    public ResponseEntity<List<ErrorResponse>> migrateDivision(@RequestPart(value = "file") MultipartFile file,
                                                HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateDivision(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("zona")
    public ResponseEntity<List<ErrorResponse>> migrateZona(@RequestPart(value = "file") MultipartFile file,
                                            HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateZona(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("area")
    public ResponseEntity<List<ErrorResponse>> migrateArea(@RequestPart(value = "file") MultipartFile file,
                                            HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateArea(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("subarea")
    public ResponseEntity<List<ErrorResponse>> migrateSubarea(@RequestPart(value = "file") MultipartFile file,
                                               HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateSubarea(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("departamento")
    public ResponseEntity<List<ErrorResponse>> migrateDepartamento(@RequestPart(value = "file") MultipartFile file,
                                                    HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateDepartamento(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("cost-centers")
    public ResponseEntity<List<ErrorResponse>> migrateCostCenters(@RequestPart(value = "file") MultipartFile file,
                                                   HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateCostCenters(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("cost-centers-org-entities-geographic")
    public ResponseEntity<List<ErrorResponse>> migrateCostCentersOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file,
                                                                        HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateCostCentersOrgEntitiesGeographic(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("cost-centers-org-entities-organizative")
    public ResponseEntity<List<ErrorResponse>> migrateCostCentersOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file,
                                                                          HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateCostCentersOrgEntitiesOrganizative(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("stores")
    public ResponseEntity<List<ErrorResponse>> migrateStores(@RequestPart(value = "file") MultipartFile file,
                                              HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateStores(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("stores-org-entities-geographic")
    public ResponseEntity<List<ErrorResponse>> migrateStoresOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file,
                                                                   HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateStoresOrgEntitiesGeographic(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("stores-org-entities-organizative")
    public ResponseEntity<List<ErrorResponse>> migrateStoresOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file,
                                                                     HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateStoresOrgEntitiesOrganizative(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("work-positions")
    public ResponseEntity<List<ErrorResponse>> migrateWorkPositions(@RequestPart(value = "file") MultipartFile file,
                                                     HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateWorkPositions(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("work-positions-details")
    public ResponseEntity<List<ErrorResponse>> migrateWorkPositionsDetails(@RequestPart(value = "file") MultipartFile file,
                                                            HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateWorkPositionsDetails(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("profiles")
    public ResponseEntity<List<ErrorResponse>> migrateProfiles(@RequestPart(value = "file") MultipartFile file,
                                                HttpServletRequest request){

        return new ResponseEntity<>(migrationService.migrateProfiles(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("profiles-groups")
    public ResponseEntity<List<ErrorResponse>> migrateProfilesGroups(@RequestPart(value = "file") MultipartFile file,
                                                      HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateProfilesGroups(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("references")
    public ResponseEntity<List<ErrorResponse>> migrateReferences(@RequestPart(value = "file") MultipartFile file,
                                                  HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateReferences(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("info-bancaria")
    public ResponseEntity<List<ErrorResponse>> migrateInfoBancaria(@RequestPart(value = "file") MultipartFile file,
                                                    HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateInfoBancaria(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("info-sueldos")
    public ResponseEntity<List<ErrorResponse>> migrateInfoSueldos(@RequestPart(value = "file") MultipartFile file,
                                                   HttpServletRequest request){
        return new ResponseEntity<>(migrationService.migrateInfoSueldos(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
    @PostMapping("/load-compensations")
    public ResponseEntity<List<ErrorResponse>> cargarCompensaciones(@RequestPart(value = "file") MultipartFile file,
                                                     HttpServletRequest request) {

        return new ResponseEntity<>(migrationService.loadCompensationsCategories(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("/load-tabs")
    public ResponseEntity<List<ErrorResponse>> loadTabs(@RequestPart(value = "file") MultipartFile file,
                                         HttpServletRequest request) {
        return new ResponseEntity<>(migrationService.loadTabs(file, request.getHeader("Authorization")), HttpStatus.OK);
    }

    @PostMapping("/load-work-position-categories")
    public ResponseEntity<List<ErrorResponse>> loadWorkPositionCategories(@RequestPart(value = "file") MultipartFile file,
                                                           HttpServletRequest request) {
        return new ResponseEntity<>(migrationService.loadWorkPositionCategories(file, request.getHeader("Authorization")) ,HttpStatus.OK);
    }
    
    @PostMapping("/load-groups")
    public ResponseEntity<List<ErrorResponse>> loadWorkGroups(@RequestPart(value = "file") MultipartFile file,
                                               HttpServletRequest request) {
        return new ResponseEntity<>(migrationService.loadGroups(file, request.getHeader("Authorization")), HttpStatus.OK);
    }
}
