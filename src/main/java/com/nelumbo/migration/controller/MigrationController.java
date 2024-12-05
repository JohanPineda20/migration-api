package com.nelumbo.migration.controller;

import com.nelumbo.migration.feign.dto.responses.UtilResponse;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;
import com.nelumbo.migration.service.MigrationService;
import jakarta.servlet.http.HttpServletRequest;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
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
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
@RestController
@RequestMapping("/migration")
@RequiredArgsConstructor
public class MigrationController {

    private static final String APPLICATION_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String ATTACHMENT_FILENAME = "attachment; filename=";

    private final MigrationService migrationService;

    @PostMapping("empresa")
    public Map<String, Object> migrateEmpresa(@RequestPart(value = "file") MultipartFile file,
                                               HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateEmpresa(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("region")
    public Map<String, Object> migrateRegion(@RequestPart(value = "file") MultipartFile file,
                                                             HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateRegion(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("division")
    public Map<String, Object> migrateDivision(@RequestPart(value = "file") MultipartFile file,
                                                HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateDivision(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("zona")
    public Map<String, Object> migrateZona(@RequestPart(value = "file") MultipartFile file,
                                            HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateZona(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("area")
    public Map<String, Object> migrateArea(@RequestPart(value = "file") MultipartFile file,
                                            HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateArea(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("subarea")
    public Map<String, Object> migrateSubarea(@RequestPart(value = "file") MultipartFile file,
                                               HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateSubarea(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("departamento")
    public Map<String, Object> migrateDepartamento(@RequestPart(value = "file") MultipartFile file,
                                                    HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateDepartamento(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("cost-centers")
    public Map<String, Object> migrateCostCenters(@RequestPart(value = "file") MultipartFile file,
                                                   HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateCostCenters(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("cost-centers-org-entities-geographic")
    public Map<String, Object> migrateCostCentersOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file,
                                                                        HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateCostCentersOrgEntitiesGeographic(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("cost-centers-org-entities-organizative")
    public Map<String, Object> migrateCostCentersOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file,
                                                                          HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateCostCentersOrgEntitiesOrganizative(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("stores")
    public Map<String, Object> migrateStores(@RequestPart(value = "file") MultipartFile file,
                                              HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateStores(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("stores-org-entities-geographic")
    public Map<String, Object> migrateStoresOrgEntitiesGeographic(@RequestPart(value = "file") MultipartFile file,
                                                                   HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateStoresOrgEntitiesGeographic(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("stores-org-entities-organizative")
    public Map<String, Object> migrateStoresOrgEntitiesOrganizative(@RequestPart(value = "file") MultipartFile file,
                                                                     HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateStoresOrgEntitiesOrganizative(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("work-positions")
    public Map<String, Object> migrateWorkPositions(@RequestPart(value = "file") MultipartFile file,
                                                     HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateWorkPositions(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("work-positions-details")
    public Map<String, Object> migrateWorkPositionsDetails(@RequestPart(value = "file") MultipartFile file,
                                                            HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateWorkPositionsDetails(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("profiles")
    public Map<String, Object> migrateProfiles(@RequestPart(value = "file") MultipartFile file,
                                                HttpServletRequest request){

        UtilResponse utilResponse = migrationService.migrateProfiles(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("references")
    public Map<String, Object> migrateReferences(@RequestPart(value = "file") MultipartFile file,
                                                  HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateReferences(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("info-bancaria")
    public Map<String, Object> migrateInfoBancaria(@RequestPart(value = "file") MultipartFile file,
                                                    HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateInfoBancaria(file, request.getHeader("Authorization"));
        return processFile(utilResponse);

    }

    @PostMapping("profiles-activation")
    public Map<String, Object> profilesDraftActivation(@RequestPart(value = "file") MultipartFile file,
                                                   HttpServletRequest request){
        UtilResponse utilResponse = migrationService.profilesDraftActivation(file, request.getHeader("Authorization"));
        return processFile(utilResponse);

    }

    @PostMapping("info-sueldos")
    public Map<String, Object> migrateInfoSueldos(@RequestPart(value = "file") MultipartFile file,
                                                   HttpServletRequest request){
        UtilResponse utilResponse = migrationService.migrateInfoSueldos(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("/load-compensations")
    public Map<String, Object> loadCompensationsCategories(@RequestPart(value = "file") MultipartFile file,
                                                     HttpServletRequest request) {

        UtilResponse utilResponse = migrationService.loadCompensationsCategories(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("/load-tabs")
    public Map<String, Object> loadTabs(@RequestPart(value = "file") MultipartFile file,
                                         HttpServletRequest request) {
        UtilResponse utilResponse = migrationService.loadTabs(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    @PostMapping("/load-work-position-categories")
    public Map<String, Object> loadWorkPositionCategories(@RequestPart(value = "file") MultipartFile file,
                                                          HttpServletRequest request) {
        UtilResponse utilResponse = migrationService.loadWorkPositionCategories(file, request.getHeader("Authorization"));
        return processFile(utilResponse);
    }

    private Map<String, Object> processFile(UtilResponse utilResponse) {
        Map<String, Object> response = new HashMap<>();
        response.put("success", utilResponse.getSuccess());
        response.put("failure", utilResponse.getFailure());

        // Convertir el archivo a base64
        String base64File = Base64.getEncoder().encodeToString(utilResponse.getByteArrayOutputStream().toByteArray());
        response.put("file", base64File);

        return response;
        /*ByteArrayResource resource = new ByteArrayResource(utilResponse.getByteArrayOutputStream().toByteArray());

        HttpHeaders headers = new HttpHeaders();
        headers.add("success", utilResponse.getSuccess().toString());
        headers.add("failure", utilResponse.getFailure().toString());
        headers.add(HttpHeaders.CONTENT_DISPOSITION, ATTACHMENT_FILENAME + "migration_errors.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .contentLength(resource.contentLength())
                .contentType(MediaType.parseMediaType(APPLICATION_EXCEL))
                .body(resource);*/
    }
}
