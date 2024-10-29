package com.nelumbo.migration.service;

import com.nelumbo.migration.exceptions.ErrorResponseException;
import com.nelumbo.migration.exceptions.NullCellException;
import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class MigrationService {

    private static final String ACTIVO_STATUS = "ACTIVO";
    private static final String INACTIVO_STATUS = "INACTIVO";
    private static final String BEARER = "Bearer ";
    private static final String MODIFIED = "modified_";
    private static final String SHEET = "Estamos con la hoja: ";
    private static final String COUNTROWS = "La cantidad de filas es: ";
    @Value("${email}")
    private String email;
    @Value("${password}")
    private String password;

    private final LoginFeign loginFeign;
    private final MigrationFeign migrationFeign;

    public void migrateEmpresa(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empresa");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 1; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 1L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empresa: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empresa: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateRegion(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("regiones");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 2; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellEmpresa = row.getCell(1);
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 2L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateDivision(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("divisiones");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 2; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellRegion = row.getCell(1);
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 2L, cellRegion.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 3L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet divisiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet divisiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateZona(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("zonas");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 2; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellDivision = row.getCell(1);
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 3L, cellDivision.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 4L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet zonas: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet zonas: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateArea(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("áreas");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 1; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 5L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet áreas: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet áreas: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateSubarea(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("subareas");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 2; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellArea = row.getCell(1);
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 6L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet subareas: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet subareas: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateDepartamento(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("departamentos");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 2; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    OrgEntityDetailRequest orgEntityDetailRequest = new OrgEntityDetailRequest();
                    Cell cellName = row.getCell(0);
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellSubarea = row.getCell(1);
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 6L, cellSubarea.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 7L);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet departamentos: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet departamentos: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateCostCenters(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("ceco");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 6; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    CostCenterRequest costCenterRequest = new CostCenterRequest();
                    Cell cellCode = row.getCell(0);
                    costCenterRequest.setCode(cellCode.getStringCellValue());
                    costCenterRequest.setDenomination(row.getCell(1).getStringCellValue());

                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(2).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(2).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken,countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(3).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(3).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(4).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(4).getStringCellValue().concat(" not found"))));

                    costCenterRequest.setCountryId(countryId);
                    costCenterRequest.setStateId(stateId);
                    costCenterRequest.setCityId(cityId);
                    Cell cellStatus = row.getCell(5);
                    long statusId = getStatusId(cellStatus);
                    costCenterRequest.setStatusId(statusId);
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            costCenterRequest.getFieldsValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createCostCenter(bearerToken, costCenterRequest);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet centro de costos: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet centro de costos: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateCostCentersOrgEntitiesGeographic(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("ceco_estructura_geografica");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long costCenterId = migrationFeign.findCostCenterByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    CostCenterDetailRequest costCenterDetailRequest = new CostCenterDetailRequest();
                    List<Long> orgEntityDetailIds = costCenterDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;

                    Cell cellRegion = row.getCell(2);
                    Cell cellDivision = row.getCell(3);
                    Cell cellZona = row.getCell(4);

                    Cell cellEmpresa = row.getCell(1);
                    if(cellEmpresa == null || cellEmpresa.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Invalid geographic structure: missing empresa");
                    }
                    Long empresaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa.getStringCellValue()).getData().getId();
                    orgEntityDetailIds.add(empresaId);
                    if (cellRegion != null && !cellRegion.getStringCellValue().isEmpty() || cellDivision != null && !cellDivision.getStringCellValue().isEmpty() || cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                        if (cellRegion != null && !cellRegion.getStringCellValue().isEmpty()) {
                            regionId = getEntityId(bearerToken, cellRegion, 2L, empresaId, "region");
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null && !cellDivision.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region and division");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing division");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona");
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_geografica: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_geografica: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateCostCentersOrgEntitiesOrganizative(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("ceco_estructura_organizativa");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long costCenterId = migrationFeign.findCostCenterByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    CostCenterDetailRequest costCenterDetailRequest = new CostCenterDetailRequest();
                    List<Long> orgEntityDetailIds = costCenterDetailRequest.getOrgEntityDetailIds();

                    Cell cellArea = row.getCell(1);
                    Cell cellSubArea = row.getCell(2);
                    Cell cellDepartamento = row.getCell(3);

                    Long subAreaId = null;
                    Long departamentoId = null;

                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Invalid geographic structure: missing area");
                    }
                    Long areaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getId();
                    orgEntityDetailIds.add(areaId);
                    if (cellSubArea != null && !cellSubArea.getStringCellValue().isEmpty() || cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) {
                        if (cellSubArea != null && !cellSubArea.getStringCellValue().isEmpty()) {
                            subAreaId = getEntityId(bearerToken, cellSubArea, 6L, areaId, "subarea");
                            orgEntityDetailIds.add(subAreaId);
                        }

                        if (cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) {
                            if (subAreaId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing subarea");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento");
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_organizativa: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_organizativa: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateStores(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sucursales");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    StoreRequest storeRequest = new StoreRequest();
                    Cell code = row.getCell(0);
                    storeRequest.setCode(code.getStringCellValue());
                    storeRequest.setDenomination(row.getCell(1).getStringCellValue());

                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(2).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(2).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken, countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(3).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(3).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(4).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(4).getStringCellValue().concat(" not found"))));

                    storeRequest.setCountryId(countryId);
                    storeRequest.setStateId(stateId);
                    storeRequest.setCityId(cityId);
                    storeRequest.setAddress(row.getCell(5) == null || row.getCell(5).getStringCellValue().isEmpty() ? "-" : row.getCell(5).getStringCellValue());
                    storeRequest.setZipcode("" + (int) row.getCell(6).getNumericCellValue());
                    storeRequest.setLatitude(row.getCell(7).getNumericCellValue());
                    storeRequest.setLongitude(row.getCell(8).getNumericCellValue());
                    storeRequest.setGeorefDistance((long) row.getCell(9).getNumericCellValue());
                    String costCenter = row.getCell(10) != null && !row.getCell(10).getStringCellValue().isEmpty()  ? row.getCell(10).getStringCellValue() : null;
                    Long costCenterId = null;
                    if(costCenter != null) {
                        costCenterId = migrationFeign.findCostCenterByCode(bearerToken, costCenter).getData().getId();
                    }
                    storeRequest.setCostCenterId(costCenterId);
                    Cell cellStatus = row.getCell(11);
                    long statusId = getStatusId(cellStatus);
                    storeRequest.setStatusId(statusId);
                    migrationFeign.createStore(bearerToken, storeRequest);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateStoresOrgEntitiesGeographic(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursales");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long storeId = migrationFeign.findStoreByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;


                    Cell cellRegion = row.getCell(13);
                    Cell cellDivision = row.getCell(14);
                    Cell cellZona = row.getCell(15);

                    Cell cellEmpresa = row.getCell(12);
                    if(cellEmpresa == null || cellEmpresa.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Invalid geographic structure: missing empresa");
                    }
                    Long empresaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa.getStringCellValue()).getData().getId();
                    orgEntityDetailIds.add(empresaId);
                    if (cellRegion != null && !cellRegion.getStringCellValue().isEmpty() || cellDivision != null && !cellDivision.getStringCellValue().isEmpty() || cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                        if (cellRegion != null && !cellRegion.getStringCellValue().isEmpty()) {
                            regionId = getEntityId(bearerToken, cellRegion, 2L, empresaId, "region");
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null && !cellDivision.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region and division");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing division");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona");
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateStoresOrgEntitiesOrganizative(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursal_estructura_organizativ");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long storeId = migrationFeign.findStoreByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Cell cellArea = row.getCell(1);
                    Cell cellSubArea = row.getCell(2);
                    Cell cellDepartamento = row.getCell(3);

                    Long subAreaId = null;
                    Long departamentoId = null;

                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Invalid geographic structure: missing area");
                    }
                    Long areaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getId();
                    orgEntityDetailIds.add(areaId);
                    if (cellSubArea != null && !cellSubArea.getStringCellValue().isEmpty() || cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) {
                        if (cellSubArea != null && !cellSubArea.getStringCellValue().isEmpty()) {
                            subAreaId = getEntityId(bearerToken, cellSubArea, 6L, areaId, "subarea");
                            orgEntityDetailIds.add(subAreaId);
                        }

                        if (cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) {
                            if (subAreaId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing subarea");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento");
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_organizativ: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_organizativ: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    private long getStatusId(Cell cellStatus) {
        long statusId = 2L;
        if (cellStatus != null && !cellStatus.getStringCellValue().isEmpty()) {
            String statusValue = cellStatus.getStringCellValue().trim().toUpperCase();
            statusId = switch (statusValue) {
                case ACTIVO_STATUS -> 1L;
                case INACTIVO_STATUS -> 2L;
                default -> throw new RuntimeException("Invalid status: " + statusValue);
            };
        }
        return statusId;
    }

    private Long getEntityId(String bearerToken, Cell cell, Long entityType, Long parentId, String entityName) {
        DefaultResponse<Page<OrgEntityResponse>> entityResponse = migrationFeign.findAllInstancesParentOrganizationEntityDetail(
                bearerToken, entityType, parentId
        );

        String name = migrationFeign.findOrgEntityDetailByName(bearerToken, entityType, cell.getStringCellValue()).getData().getName();

        return entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(name))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(cell.getStringCellValue()).concat(" not found")));
    }

    public void migrateWorkPositions(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 14; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
                String encabezado = rowEncabezados.getCell(i).getStringCellValue();
                fieldValues.put(encabezado, i);
            }

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    WorkPositionRequest workPositionRequest = new WorkPositionRequest();
                    Cell code = row.getCell(0);
                    workPositionRequest.setCode(code.getStringCellValue());
                    workPositionRequest.setDenomination(row.getCell(1).getStringCellValue());
                    workPositionRequest.setAuthorizedStaff((long)row.getCell(2).getNumericCellValue());

                    Long workPosCatId = migrationFeign.findWorkPosCategoryByCode(bearerToken, row.getCell(3).getStringCellValue()).getData().getId();
                    workPositionRequest.setWorkPosCatId(workPosCatId);

                    Long storeId = migrationFeign.findStoreByCode(bearerToken, row.getCell(4).getStringCellValue()).getData().getId();
                    workPositionRequest.setStoreId(storeId);

                    String costCenter = row.getCell(5) != null && !row.getCell(5).getStringCellValue().isEmpty() ? row.getCell(5).getStringCellValue() : null;
                    Long costCenterId = null;
                    if(costCenter != null) {
                        costCenterId = migrationFeign.findCostCenterByCode(bearerToken, costCenter).getData().getId();
                    }
                    workPositionRequest.setCostCenterId(costCenterId);
                    Cell cellStatus = row.getCell(6);
                    long statusId = getStatusId(cellStatus);
                    workPositionRequest.setStatusId(statusId);

                    Cell cellArea = row.getCell(7);
                    Cell cellSubarea = row.getCell(8);
                    Cell cellDepartamento = row.getCell(9);
                    Long storeOrganizativeId = null;

                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()) throw new RuntimeException("Area is required");

                    DefaultResponse<StoreDetailResponse> storeDetailResponse = migrationFeign.findAllStoresDetails(bearerToken, storeId);
                    //Obtener las estructuras organizativas de la sucursal cuya area sea igual a cellArea
                    String area = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getName();
                    List<OrgEntDetailResponse> areasFiltradas = storeDetailResponse.getData().getStructuresByType().stream()
                            .flatMap(structureType -> structureType.getDetails().stream())
                            .filter(detail -> detail.getStructures().stream().anyMatch(structure -> area.equalsIgnoreCase(structure.getName()) && structure.getOrgEntity().getId() == 5L))
                            .toList();
                    //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa area
                    if (areasFiltradas.isEmpty()) throw new RuntimeException("Area ".concat(cellArea.getStringCellValue()).concat(" not found. It must match the structure of the store."));

                    if (cellSubarea != null && !cellSubarea.getStringCellValue().isEmpty()) {
                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas tienen el subarea
                        String subArea = migrationFeign.findOrgEntityDetailByName(bearerToken, 6L, cellSubarea.getStringCellValue()).getData().getName();
                        List<OrgEntDetailResponse> areasFiltradasConSubarea = areasFiltradas.stream()
                                .filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> subArea.equalsIgnoreCase(child.getName()) && child.getOrgEntity().getId() == 6L)))
                                .toList();
                        //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa subarea
                        if (areasFiltradasConSubarea.isEmpty()) throw new RuntimeException("Subarea ".concat(cellSubarea.getStringCellValue()).concat(" not found. It must match the structure of the store."));

                        if (cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) {
                            //Una vez encontradas las estructuras organizativas que tienen ese area-subarea, buscar cual de ellas tienen el departamento
                            String departamento = migrationFeign.findOrgEntityDetailByName(bearerToken, 7L, cellDepartamento.getStringCellValue()).getData().getName();
                            Optional<OrgEntDetailResponse> areaConSubareaYDepartamento = areasFiltradasConSubarea.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> subArea.equalsIgnoreCase(child.getName()) && child.getOrgEntity().getId() == 6L && child.getChildren().stream().anyMatch(child2 -> child2.getName().equalsIgnoreCase(departamento) && child2.getOrgEntity().getId() == 7L))))
                                    .findFirst();

                            //Si el optional es vacio es porque ningun area-subarea tiene ese departamento
                            if (areaConSubareaYDepartamento.isEmpty()) throw new RuntimeException("Departamento ".concat(cellDepartamento.getStringCellValue()).concat(" not found. It must match the structure of the store."));
                            storeOrganizativeId = areaConSubareaYDepartamento.get().getId();
                        }
                        else {
                            //Una vez encontradas las estructuras organizativas que tienen ese area-subarea, buscar cual de ellas no tiene departamento
                            Optional<OrgEntDetailResponse> areaConSubareaSinDepartamento = areasFiltradasConSubarea.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren().stream().anyMatch(child -> child.getChildren() == null || child.getChildren().isEmpty() || child.getChildren().get(0) == null)))
                                    .findFirst();
                            //Si el optional es vacio es porque todas las area-subarea tienen un departamento y se necesita que en el excel se envíe el departamento para buscarlo
                            if(areaConSubareaSinDepartamento.isEmpty()) throw new RuntimeException("A Departamento is required");
                            storeOrganizativeId = areaConSubareaSinDepartamento.get().getId();
                        }
                    }
                    else  {
                        //Si el excel tiene un departamento y no tiene un subarea, entonces está mal la estructura, falta el subarea
                        if(cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) throw new RuntimeException("A Subarea is required");

                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas no tiene subarea
                        Optional<OrgEntDetailResponse> areaSinSubarea = areasFiltradas.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() == null || structure.getChildren().isEmpty() || structure.getChildren().get(0) == null))
                                .findFirst();
                        //Si el optional es vacio es porque todas las areas tienen un subarea y se necesita que en el excel se envíe el subarea para buscarlo
                        if(areaSinSubarea.isEmpty()) throw new RuntimeException("A Subarea is required");
                        storeOrganizativeId = areaSinSubarea.get().getId();
                    }
                    workPositionRequest.setStoreOrganizativeId(storeOrganizativeId);

                    fieldValues.forEach((nameColumn, position) -> {
                        Cell cell = row.getCell(position);
                        if (cell == null) {
                            workPositionRequest.getFieldsValues().put(nameColumn, null);
                        } else {
                            switch (cell.getCellType()) {
                                case STRING:
                                    workPositionRequest.getFieldsValues().put(nameColumn, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        LocalDate formattedDate = cell.getDateCellValue().toInstant()
                                                .atZone(ZoneId.systemDefault())
                                                .toLocalDate();
                                        DateTimeFormatter pattern = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                                        workPositionRequest.getFieldsValues().put(nameColumn, formattedDate.format(pattern));
                                    } else {
                                        workPositionRequest.getFieldsValues().put(nameColumn, (long) cell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    workPositionRequest.getFieldsValues().put(nameColumn, cell.getBooleanCellValue());
                                    break;
                                default:
                                    workPositionRequest.getFieldsValues().put(nameColumn, null);
                                    break;
                            }
                        }
                    });

                    migrationFeign.createWorkPosition(bearerToken, workPositionRequest);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateWorkPositionsDetails(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long workPositionId = migrationFeign.findWorkPositionByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getWorkPosition().getId();
                    String compCategory = row.getCell(10) != null && !row.getCell(10).getStringCellValue().isEmpty() ? row.getCell(10).getStringCellValue() : null;
                    Long compCategoryId = null;
                    if(compCategory != null){
                        compCategoryId = migrationFeign.findCompCategoryByCode(bearerToken, compCategory).getData().getId();
                    }
                    String compTab = row.getCell(11) != null && !row.getCell(11).getStringCellValue().isEmpty() ? row.getCell(11).getStringCellValue() : null;
                    Long compTabId = null;
                    if(compTab != null){
                        compTabId = migrationFeign.findCompTabByCode(bearerToken, compTab).getData().getId();
                    }
                    String managerWorkPosition = row.getCell(12) != null && !row.getCell(12).getStringCellValue().isEmpty() ? row.getCell(12).getStringCellValue() : null;
                    Long managerWorkPositionId = null;
                    if(managerWorkPosition != null){
                        managerWorkPositionId = migrationFeign.findWorkPositionByCode(bearerToken, managerWorkPosition).getData().getWorkPosition().getId();
                    }

                    Long authorizedSalary = row.getCell(13) != null && row.getCell(13).getNumericCellValue() != 0 ? Math.round(row.getCell(13).getNumericCellValue()) : null;

                    if(!(compCategoryId == null && compTabId == null && managerWorkPositionId == null && authorizedSalary == null)){
                        WorkPositionUpdateRequest wPUReq = WorkPositionUpdateRequest.builder()
                                .compCategoryId(compCategoryId)
                                .compTabId(compTabId)
                                .orgManagerId(managerWorkPositionId)
                                .approvalManagerId(managerWorkPositionId)
                                .minSalary(authorizedSalary)
                                .build();
                        migrationFeign.updateWorkPosition(bearerToken, wPUReq, workPositionId);
                    }
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateProfiles(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empleados");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    ProfileRequest profileRequest = new ProfileRequest();
                    List<ProfileSecValueRequest> profileSecValueRequestList = new ArrayList<>();
                    Cell clave = row.getCell(0);
                    ProfileSecValueRequest informacionPersonal = new ProfileSecValueRequest();
                    informacionPersonal.setKeyword("PSPI01");
                    Map<String, Object> informacionPersonalValues = informacionPersonal.getFieldsValues();
                    informacionPersonalValues.put("Primer Nombre", row.getCell(1).getStringCellValue());
                    informacionPersonalValues.put("Primer Apellido", row.getCell(2).getStringCellValue());
                    informacionPersonalValues.put("Segundo Apellido", row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "");
                    informacionPersonalValues.put("Sexo", row.getCell(4).getStringCellValue());
                    DateTimeFormatter formatters = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    //LocalDate.parse(row.getCell(9).getStringCellValue(), formatters);
                    // row.getCell(9).getStringCellValue()
                    LocalDate hiredDate =  row.getCell(9).getDateCellValue().toInstant()
                            .atZone(ZoneId.systemDefault())
                            .toLocalDate();
                    informacionPersonalValues.put("Fecha de contratación", hiredDate.format(formatters));
                    informacionPersonalValues.put("Clave MPRO", clave != null ? clave.getStringCellValue() : "");

                    ProfileSecValueRequest informacionBiografica = new ProfileSecValueRequest();
                    informacionBiografica.setKeyword("PSBI02");
                    Map<String, Object> informacionBiograficaValues = informacionBiografica.getFieldsValues();
                    //LocalDate.parse(row.getCell(8).getStringCellValue(), formatters);
                    //row.getCell(8).getStringCellValue()
                    LocalDate birthDate =  row.getCell(8).getDateCellValue().toInstant()
                            .atZone(ZoneId.systemDefault())
                            .toLocalDate();
                    informacionBiograficaValues.put("Fecha de nacimiento", birthDate.format(formatters));

                    ProfileSecValueRequest datosPersonales = new ProfileSecValueRequest();
                    datosPersonales.setKeyword("PSPD03");
                    Map<String, Object> datosPersonalesValues = datosPersonales.getFieldsValues();
                    datosPersonalesValues.put("RFC", row.getCell(5).getStringCellValue());
                    datosPersonalesValues.put("CURP", row.getCell(6).getStringCellValue());
                    datosPersonalesValues.put("NSS", row.getCell(7).getStringCellValue());

                    ProfileSecValueRequest direccion = new ProfileSecValueRequest();
                    direccion.setKeyword("PSAS05");
                    Map<String, Object> direccionValues = direccion.getFieldsValues();
                    direccionValues.put("Dirección", row.getCell(12).getStringCellValue());
                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    CountryResponse paisResidencia = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(13).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(13).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken, paisResidencia.getId());
                    CountryResponse estadoResidencia = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(14).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(14).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, paisResidencia.getId(), estadoResidencia.getId());
                    CountryResponse ciudadResidencia = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(15).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(15).getStringCellValue().concat(" not found"))));
                    direccionValues.put("Lugar de Residencia", Arrays.asList(paisResidencia, estadoResidencia, ciudadResidencia));

                    ProfileSecValueRequest contacto = new ProfileSecValueRequest();
                    contacto.setKeyword("PSCI06");
                    Map<String, Object> contactoValues = contacto.getFieldsValues();
                    contactoValues.put("Email Personal", row.getCell(10).getStringCellValue());
                    contactoValues.put("Celular personal", (row.getCell(11) == null) ? "" : row.getCell(11).getStringCellValue());

                    profileSecValueRequestList.add(informacionPersonal);
                    profileSecValueRequestList.add(informacionBiografica);
                    profileSecValueRequestList.add(datosPersonales);
                    profileSecValueRequestList.add(direccion);
                    profileSecValueRequestList.add(contacto);

                    profileRequest.setSectionValues(profileSecValueRequestList);

                    Long workPositionId = migrationFeign.findWorkPositionByCode(bearerToken, row.getCell(16).getStringCellValue()).getData().getWorkPosition().getId();

                    profileRequest.setWorkPositionId(workPositionId);
                    migrationFeign.createProfile(bearerToken, profileRequest);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getError().getErrors().getFields());
                    
                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateProfilesGroups(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empleados");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    String group = row.getCell(27) != null ? row.getCell(27).getStringCellValue() : null;
                    if(group != null) {

                        Long idGroup = migrationFeign.findGroupByName(bearerToken, group).getData().getId();

                        Set<Long> idProfiles = new HashSet<>();
                        Set<Long> idGroups = new HashSet<>();
                        idProfiles.add(profileId);
                        idGroups.add(idGroup);

                        GroupsProfRequest groupsProfRequest = new GroupsProfRequest();
                        groupsProfRequest.setProfileIds(idProfiles);
                        groupsProfRequest.setGroupIds(idGroups);
                        groupsProfRequest.setTemporal(false);
                        groupsProfRequest.setAllProfiles(false);

                        this.migrationFeign.createGroupsAssigments(bearerToken, groupsProfRequest);
                    }
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getError().getErrors().getFields());

                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateReferences(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("referencias");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO is required");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, cellClaveMPRO.getStringCellValue()).getData().getId();

                    ProfileSecValueRequest references = new ProfileSecValueRequest();
                    references.setKeyword("PSRF16");
                    Map<String, Object> referencesValues = references.getFieldsValues();
                    Cell cellNombre = row.getCell(1);
                    Cell cellTelefono = row.getCell(2);
                    if(cellNombre == null || cellNombre.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Name is required");
                    }
                    if(cellTelefono == null || cellTelefono.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Phone is required");
                    }
                    referencesValues.put("Nombre", cellNombre.getStringCellValue());
                    referencesValues.put("Teléfono", cellTelefono.getStringCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, references);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet referencias: " + e.getError().getErrors().getFields());

                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet referencias: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateInfoBancaria(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("informacion bancaria");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO is required");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, cellClaveMPRO.getStringCellValue()).getData().getId();

                    ProfileSecValueRequest informacionPago = new ProfileSecValueRequest();
                    informacionPago.setKeyword("PSPM14");
                    Map<String, Object> informacionPagoValues = informacionPago.getFieldsValues();
                    Cell cellBanco = row.getCell(1);
                    Cell cellCuenta = row.getCell(2);
                    Cell cellClabe = row.getCell(3);
                    Cell cellTitular = row.getCell(4);
                    if(cellBanco == null || cellBanco.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Banco is required");
                    }
                    if(cellCuenta == null || cellCuenta.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Cuenta bancaria is required");
                    }
                    if(cellClabe == null || cellClabe.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clabe interbancaria is required");
                    }
                    if(cellTitular == null || cellTitular.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Titular de la cuenta is required");
                    }

                    informacionPagoValues.put("Banco", cellBanco.getStringCellValue().toUpperCase());
                    informacionPagoValues.put("Cuenta bancaria", cellCuenta.getStringCellValue());
                    informacionPagoValues.put("Clabe interbancaria", cellClabe.getStringCellValue());
                    informacionPagoValues.put("Titular de la cuenta", cellTitular.getStringCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, informacionPago);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet informacion bancaria: " + e.getError().getErrors().getFields());

                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet informacion bancaria: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateInfoSueldos(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sueldos");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO is required");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, cellClaveMPRO.getStringCellValue()).getData().getId();

                    ProfileSecValueRequest payrollInformation = new ProfileSecValueRequest();
                    payrollInformation.setKeyword("PSPN11");
                    Map<String, Object> payrollInformationValues = payrollInformation.getFieldsValues();
                    Cell cellSueldoMensual = row.getCell(1);
                    Cell cellSueldoDiario = row.getCell(2);
                    if(cellSueldoMensual == null || cellSueldoMensual.getNumericCellValue() == 0) {
                        throw new RuntimeException("Salario mensual is required");
                    }
                    if(cellSueldoDiario == null || cellSueldoMensual.getNumericCellValue() == 0) {
                        throw new RuntimeException("Sueldo diario is required");
                    }

                    payrollInformationValues.put("Salario mensual", cellSueldoMensual.getNumericCellValue());
                    payrollInformationValues.put("Sueldo diario", cellSueldoDiario.getNumericCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, payrollInformation);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sueldos: " + e.getError().getErrors().getFields());

                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sueldos: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public File loadCompensationsCategories(MultipartFile file) {

        String bearerToken = this.getBearerToken();

        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("categorias de puesto");

            logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new HashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if (columnName == null) {
                    cellCode = i;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("CODIGO")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("DENOMINACION")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("ESTATUS")) {
                    cellStatus = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null || cellStatus == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells());
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("CODIGO / DENOMINACION / ESTATUS column does not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("CODIGO / DENOMINACION / ESTATUS column does not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String code = row.getCell(cellCode).getStringCellValue();
                    String denomination = row.getCell(cellDenomination).getStringCellValue();
                    Map<String, Object> fieldsValues = new HashMap<>();

                    fieldsExcel.forEach((nameColumn, position) -> {
                        Cell cell = row.getCell(position);
                        if (cell == null) {
                            fieldsValues.put(nameColumn, null);
                        } else {
                            switch (cell.getCellType()) {
                                case STRING:
                                    fieldsValues.put(nameColumn, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        fieldsValues.put(nameColumn, cell.getDateCellValue());
                                    } else {
                                        if(Objects.equals(nameColumn, "Fondo de ahorro")){
                                            fieldsValues.put(nameColumn, cell.getNumericCellValue());
                                        }
                                        else {
                                            fieldsValues.put(nameColumn, (long) cell.getNumericCellValue());
                                        }
                                    }
                                    break;
                                case BOOLEAN:
                                    fieldsValues.put(nameColumn, cell.getBooleanCellValue());
                                    break;
                                default:
                                    fieldsValues.put(nameColumn, null);
                                    break;
                            }
                        }
                    });

                    Cell cellStatus2 = row.getCell(cellStatus);
                    long statusId = getStatusId(cellStatus2);

                    CompCategoriesRequest compCategories = new CompCategoriesRequest();
                    compCategories.setCode(code);
                    compCategories.setDenomination(denomination);
                    compCategories.setFieldsValues(fieldsValues);
                    compCategories.setStatusId(statusId);

                    migrationFeign.createCompensationCategories(bearerToken, compCategories);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet categorias de puesto: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet categorias de puesto: " + e.getMessage());
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                }
            }
            // Archivo modificado para devolver
            modifiedFile = this.createModifiedWorkbook(workbook, file);
        } catch (Exception e) {
            this.logProcessingExcelFile(e);
        }
        return modifiedFile;
    }

    public File loadTabs(MultipartFile file) {

        String bearerToken = this.getBearerToken();

        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("tabuladores");

            this.logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new HashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if (columnName == null) {
                    cellCode = i;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("NIVEL MACROPAY")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("POSICION")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("ESTATUS")) {
                    cellStatus = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null || cellStatus == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("NIVEL MACROPAY / POSICION / ESTATUS column does not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("NIVEL MACROPAY / POSICION / ESTATUS column does not exist");
            }

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String code = row.getCell(cellCode).getStringCellValue();
                    String denomination = row.getCell(cellDenomination).getStringCellValue();
                    Map<String, Object> fieldsValues = new HashMap<>();

                    fieldsExcel.forEach((nameColumn, position) -> {
                        Cell cell = row.getCell(position);
                        if (cell == null) {
                            fieldsValues.put(nameColumn, null);
                        } else {
                            switch (cell.getCellType()) {
                                case STRING:
                                    fieldsValues.put(nameColumn, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        fieldsValues.put(nameColumn, cell.getDateCellValue());
                                    } else {
                                        fieldsValues.put(nameColumn, (long) cell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    fieldsValues.put(nameColumn, cell.getBooleanCellValue());
                                    break;
                                default:
                                    fieldsValues.put(nameColumn, null);
                                    break;
                            }
                        }
                    });

                    Cell cellStatus2 = row.getCell(cellStatus);
                    long statusId = getStatusId(cellStatus2);

                    TabsRequest tabsRequest = new TabsRequest();
                    tabsRequest.setCode(code);
                    tabsRequest.setDenomination(denomination);
                    tabsRequest.setMinAuthorizedSalary(0L);
                    tabsRequest.setMaxAuthorizedSalary(0L);
                    tabsRequest.setStatusId(statusId);
                    tabsRequest.setFieldsValues(fieldsValues);

                    migrationFeign.createTab(bearerToken, tabsRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet tabuladores: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet tabuladores: " + e.getMessage());
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                }
            }

            modifiedFile = this.createModifiedWorkbook(workbook, file);
        } catch (Exception e) {
            this.logProcessingExcelFile(e);
        }
        return modifiedFile;
    }

    public File loadWorkPositionCategories(MultipartFile file) {

        String bearerToken = this.getBearerToken();

        // Archivo modificado para devolver
        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("puestos");

            this.logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new HashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if(columnName == null) {
                    continue;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("CODIGO")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("DENOMINACION")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("ESTATUS")) {
                    cellStatus = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null || cellStatus == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("CODIGO / DENOMINACION / ESTATUS column does not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("CODIGO / DENOMINACION / ESTATUS column does not exist");
            }

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String code = row.getCell(cellCode).getStringCellValue();
                    String denomination = row.getCell(cellDenomination).getStringCellValue();
                    Map<String, Object> fieldsValues = new HashMap<>();

                    fieldsExcel.forEach((nameColumn, position) -> {
                        Cell cell = row.getCell(position);
                        log.info(nameColumn);
                        if (cell == null) {
                            fieldsValues.put(nameColumn, null);
                        } else {
                            switch (cell.getCellType()) {
                                case STRING:
                                    fieldsValues.put(nameColumn, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        fieldsValues.put(nameColumn, cell.getDateCellValue());
                                    } else {
                                        fieldsValues.put(nameColumn, (long) cell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    fieldsValues.put(nameColumn, cell.getBooleanCellValue());
                                    break;
                                default:
                                    fieldsValues.put(nameColumn, null);
                                    break;
                            }
                        }
                    });

                    Cell cellStatus2 = row.getCell(cellStatus);
                    long statusId = getStatusId(cellStatus2);

                    WorkPositionCategoryRequest workPositionCategoryRequest = new WorkPositionCategoryRequest();
                    workPositionCategoryRequest.setCode(code);
                    workPositionCategoryRequest.setDenomination(denomination);
                    workPositionCategoryRequest.setFieldsValues(fieldsValues);
                    workPositionCategoryRequest.setStatusId(statusId);

                    migrationFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getMessage());
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                }
            }
            modifiedFile = this.createModifiedWorkbook(workbook, file);
        } catch (Exception e) {
            this.logProcessingExcelFile(e);
        }
        return modifiedFile;
    }
    
    public void loadGroups(MultipartFile file) {

        String bearerToken = this.getBearerToken();

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheet("grupos");

            this.logSheetNameNumberOfRows(sheet);

            Row rowNames = sheet.getRow(0);
            Integer cellName = null;
            Integer cellDescription = null;


            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                Cell columnName = rowNames.getCell(i);

                if(columnName == null) {
                    continue;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("nombre")) {
                    cellName = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("descripcion")) {
                    cellDescription = i;
                }
            }

            if(cellName == null || cellDescription == null) {
                throw new NullCellException("name / description column do not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);

                    if(row.getCell(cellName) == null) {
                        throw new NullCellException("name cell can not be null");
                    }

                    String name = (row.getCell(cellName).getStringCellValue()).trim();
                    String description = (row.getCell(cellDescription) == null) ? null : (row.getCell(cellDescription).getStringCellValue()).trim();

                    log.info("Group with name: " + name + "\ndescription: " + description);

                    // Preparamos el objeto que irá en el body
                    GroupsRequest groupsRequest = new GroupsRequest();
                    groupsRequest.setName(name);
                    groupsRequest.setDescription(description);

                    // Realizamos la petición
                    this.migrationFeign.createGroups(bearerToken, groupsRequest);
                } catch(ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet grupos: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet grupos: " + e.getMessage());
                }
            }

        } catch(ErrorResponseException e) {
            log.error("Error searching groups, Description: " + e.getError().getErrors().getDescription() 
                        + "\n Fields: " + e.getError().getErrors().getFields().toString());
        } catch (Exception e) {
            this.logProcessingExcelFile(e);
        }
    }

    private String getBearerToken() {
        // Realizamos el login para obtener un token
        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(password);
        return  BEARER.concat(loginFeign.login(loginRequest).getData().getToken());
    }

    private void logSheetNameNumberOfRows(Sheet sheet) {
        //Imprimimos el nombre de la hoja
        log.info(SHEET + sheet.getSheetName());

        //Imprimimos el numeros de filas en la hoja
        log.info(COUNTROWS + sheet.getPhysicalNumberOfRows());
    }

    private CellStyle greenCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        XSSFColor greenColor = new XSSFColor(java.awt.Color.GREEN, null);
        ((XSSFCellStyle) cellStyle).setFillForegroundColor(greenColor);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }
    
    private CellStyle redCellStyle(Workbook workbook) {
        CellStyle redCellStyle = workbook.createCellStyle();
        XSSFColor redColor = new XSSFColor(java.awt.Color.RED, null);
        ((XSSFCellStyle) redCellStyle).setFillForegroundColor(redColor);
        redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return redCellStyle;
    }

    private void logProcessingExcelFile(Exception e) {
        log.error("Error processing Excel file: " + e.getMessage());
    }

    private void agregarCeldaError(Row row, String message) {
        // Agregar celda con el mensaje de error en la fila que falló
        Cell errorCell = row.createCell(row.getPhysicalNumberOfCells() + 1);
        errorCell.setCellValue("Error: " + message);
    }
    
    private void agregarExcetionFeign(Row row, List<String> fields) {
        // Agregar celda con el mensaje de error en la fila que falló
        Cell errorCell = row.createCell(row.getPhysicalNumberOfCells() + 1);
        StringBuilder errores = new StringBuilder();
        for (String f : fields) {
            errores = errores.append(f).append(" ");
        }

        errorCell.setCellValue("Error: " + errores.toString());
    }

    private File createModifiedWorkbook(Workbook workbook, MultipartFile file) {
        // Archivo modificado para devolver
        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Escribir el workbook modificado de nuevo en el archivo original
        try (FileOutputStream fileOut = new FileOutputStream(modifiedFile)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            log.error("Error writing modified Excel file: " + e.getMessage());
        }
        return modifiedFile;
    }
}
