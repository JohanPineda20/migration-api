package com.nelumbo.migration.service;

import com.nelumbo.migration.exceptions.*;
import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import com.nelumbo.migration.feign.dto.responses.error.ErrorDetailResponse;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;

import feign.FeignException;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class MigrationService {

    private static final String ACTIVO_STATUS = "ACTIVO";
    private static final String INACTIVO_STATUS = "INACTIVO";
    private static final String SHEET = "Estamos con la hoja: ";
    private static final String COUNTROWS = "La cantidad de filas es: ";

    private final MigrationFeign migrationFeign;

    public UtilResponse migrateEmpresa(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empresa");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) -> {
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 1L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet empresa: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet empresa: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateRegion(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("regiones");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
            Map<String, Long> parentIdCache = new ConcurrentHashMap<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    String cellEmpresa = getCellValueAsString(row.getCell(1));
                    if(cellEmpresa.isEmpty()) {
                        throw new RuntimeException("Empresa es requerido");
                    }
                    Long parentId = parentIdCache.computeIfAbsent(cellEmpresa, empresa ->
                            migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, empresa).getData().getId()
                    );
                    orgEntityDetailRequest.setParentId(parentId);

                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 2L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateDivision(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("divisiones");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
            Map<String, Long> parentIdCache = new ConcurrentHashMap<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    String cellRegion = getCellValueAsString(row.getCell(1));
                    if(cellRegion.isEmpty()) {
                        throw new RuntimeException("Region es requerido");
                    }
                    Long parentId = parentIdCache.computeIfAbsent(cellRegion, region ->
                            migrationFeign.findOrgEntityDetailByName(bearerToken, 2L, region).getData().getId()
                    );
                    orgEntityDetailRequest.setParentId(parentId);

                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 3L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet divisiones: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet divisiones: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateZona(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("zonas");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
            Map<String, Long> parentIdCache = new ConcurrentHashMap<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    String cellDivision = getCellValueAsString(row.getCell(1));
                    if(cellDivision.isEmpty()) {
                        throw new RuntimeException("Division es requerido");
                    }
                    Long parentId = parentIdCache.computeIfAbsent(cellDivision, division ->
                            migrationFeign.findOrgEntityDetailByName(bearerToken, 3L, division).getData().getId()
                    );
                    orgEntityDetailRequest.setParentId(parentId);

                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 4L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet zonas: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet zonas: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateArea(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("áreas");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 5L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet áreas: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet áreas: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateSubarea(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("subareas");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
            Map<String, Long> parentIdCache = new ConcurrentHashMap<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    String cellArea = getCellValueAsString(row.getCell(1));
                    if(cellArea.isEmpty()) {
                        throw new RuntimeException("Area es requerida");
                    }
                    Long parentId = parentIdCache.computeIfAbsent(cellArea, area ->
                            migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, area).getData().getId()
                    );
                    orgEntityDetailRequest.setParentId(parentId);

                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 6L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet subareas: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet subareas: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateDepartamento(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("departamentos");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);
            Map<String, Long> parentIdCache = new ConcurrentHashMap<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    String cellSubarea = getCellValueAsString(row.getCell(1));
                    if(cellSubarea.isEmpty()) {
                        throw new RuntimeException("Subarea es requerida");
                    }
                    Long parentId = parentIdCache.computeIfAbsent(cellSubarea, subarea ->
                            migrationFeign.findOrgEntityDetailByName(bearerToken, 6L, subarea).getData().getId()
                    );
                    orgEntityDetailRequest.setParentId(parentId);

                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 7L);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 1);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet departamentos: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet departamentos: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateCostCenters(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("ceco");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> countriesCache = migrationFeign.findAll(bearerToken).getData().stream()
                    .collect(Collectors.toMap(
                            country -> country.getName().toLowerCase(), // Convertir a minúsculas
                            CountryResponse::getId
                    ));
            Map<Long, Map<String, Long>> statesCache = new ConcurrentHashMap<>();
            Map<Long, Map<Long, Map<String, Long>>> citiesCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);
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
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Codigo es requerido");
                    }
                    costCenterRequest.setCode(cellCode);
                    Cell cellDenomination = row.getCell(1);
                    if(cellDenomination == null || cellDenomination.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Denominacion es requerida");
                    }
                    costCenterRequest.setDenomination(cellDenomination.getStringCellValue());

                    String cellCountry = getCellValueAsString(row.getCell(2));
                    if (cellCountry.isEmpty()) {
                        throw new RuntimeException("Pais es requerido");
                    }
                    String cellState = getCellValueAsString(row.getCell(3));
                    if (cellState.isEmpty()) {
                        throw new RuntimeException("Estado es requerido");
                    }
                    String cellCity = getCellValueAsString(row.getCell(4));
                    if (cellCity.isEmpty()) {
                        throw new RuntimeException("Municipio es requerido");
                    }

                    Long countryId = countriesCache.get(cellCountry.toLowerCase());
                    if (countryId == null) throw new RuntimeException("Pais no encontrado: " + cellCountry);

                    synchronized (statesCache) {
                        statesCache.computeIfAbsent(countryId, id -> migrationFeign.findAllStatesByCountryId(bearerToken, id).getData().stream()
                                .collect(Collectors.toMap(
                                        state -> state.getName().toLowerCase(), // Convertir a minúsculas
                                        CountryResponse::getId
                                )));
                    }

                    Long stateId = statesCache.get(countryId).get(cellState.toLowerCase());
                    if (stateId == null) throw new RuntimeException("Estado no encontrado: " + cellState);

                    synchronized (citiesCache) {
                        citiesCache.computeIfAbsent(countryId, id -> new HashMap<>())
                                .computeIfAbsent(stateId, id -> migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId).getData().stream()
                                        .collect(Collectors.toMap(
                                                city -> city.getName().toLowerCase(), // Convertir a minúsculas
                                                CountryResponse::getId
                                        )));
                    }
                    Long cityId = citiesCache.get(countryId).get(stateId).get(cellCity.toLowerCase());
                    if (cityId == null) throw new RuntimeException("Municipio no encontrado: " + cellCity);

                    costCenterRequest.setCountryId(countryId);
                    costCenterRequest.setStateId(stateId);
                    costCenterRequest.setCityId(cityId);
                    Cell cellStatus = row.getCell(5);
                    long statusId = getStatusId(cellStatus);
                    costCenterRequest.setStatusId(statusId);
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            costCenterRequest.getFieldsValues().put(name, getCellValueAsString(cell));
                        }
                    });

                    migrationFeign.createCostCenter(bearerToken, costCenterRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateCostCentersOrgEntitiesGeographic(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("ceco_estructura_geografica");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> costCenterCache = new ConcurrentHashMap<>();
            Map<String, Long> entityCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Codigo del centro de costos es requerido");
                    }
                    Long costCenterId = costCenterCache.computeIfAbsent(cellCode, code -> migrationFeign.findCostCenterByCode(bearerToken, code).getData().getId());

                    CostCenterDetailRequest costCenterDetailRequest = new CostCenterDetailRequest();
                    List<Long> orgEntityDetailIds = costCenterDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;

                    String cellEmpresa = getCellValueAsString(row.getCell(1));
                    String cellRegion = getCellValueAsString(row.getCell(2));
                    String cellDivision = getCellValueAsString(row.getCell(3));
                    String cellZona = getCellValueAsString(row.getCell(4));

                    if(cellEmpresa.isEmpty()) {
                        throw new RuntimeException("Empresa es requerida");
                    }
                    String cacheKey = 1L + cellEmpresa;
                    Long empresaId;
                    if(entityCache.containsKey(cacheKey)) {
                        empresaId = entityCache.get(cacheKey);
                    } else {
                        empresaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa).getData().getId();
                        entityCache.put(cacheKey, empresaId);
                    }
                    orgEntityDetailIds.add(empresaId);
                    if (cellRegion != null && !cellRegion.isEmpty() || cellDivision != null && !cellDivision.isEmpty() || cellZona != null && !cellZona.isEmpty()) {
                        if (cellRegion != null && !cellRegion.isEmpty()) {
                            regionId = getEntityId(bearerToken, cellRegion, 2L, empresaId, "region", entityCache);
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null && !cellDivision.isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region es requerido");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division", entityCache);
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region y Division son requeridos");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Division es requerido");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona", entityCache);
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_geografica: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_geografica: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateCostCentersOrgEntitiesOrganizative(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("ceco_estructura_organizativa");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> costCenterCache = new ConcurrentHashMap<>();
            Map<String, Long> entityCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Codigo del centro de costos es requerido");
                    }
                    Long costCenterId = costCenterCache.computeIfAbsent(cellCode, code -> migrationFeign.findCostCenterByCode(bearerToken, code).getData().getId());

                    CostCenterDetailRequest costCenterDetailRequest = new CostCenterDetailRequest();
                    List<Long> orgEntityDetailIds = costCenterDetailRequest.getOrgEntityDetailIds();

                    String cellArea = getCellValueAsString(row.getCell(1));
                    String cellSubArea = getCellValueAsString(row.getCell(2));
                    String cellDepartamento = getCellValueAsString(row.getCell(3));

                    Long subAreaId = null;
                    Long departamentoId = null;

                    if(cellArea.isEmpty()) {
                        throw new RuntimeException("Area es requerida");
                    }
                    String cacheKey = 5L + cellArea;
                    Long areaId;
                    if(entityCache.containsKey(cacheKey)) {
                        areaId = entityCache.get(cacheKey);
                    } else {
                        areaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea).getData().getId();
                        entityCache.put(cacheKey, areaId);
                    }
                    orgEntityDetailIds.add(areaId);
                    if (cellSubArea != null && !cellSubArea.isEmpty() || cellDepartamento != null && !cellDepartamento.isEmpty()) {
                        if (cellSubArea != null && !cellSubArea.isEmpty()) {
                            subAreaId = getEntityId(bearerToken, cellSubArea, 6L, areaId, "subarea", entityCache);
                            orgEntityDetailIds.add(subAreaId);
                        }

                        if (cellDepartamento != null && !cellDepartamento.isEmpty()) {
                            if (subAreaId == null) {
                                throw new RuntimeException("Subarea es requerida");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento", entityCache);
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_organizativa: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet ceco_estructura_organizativa: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateStores(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sucursales");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> costCenterCache = new ConcurrentHashMap<>();
            Map<String, Long> countriesCache = migrationFeign.findAll(bearerToken).getData().stream()
                    .collect(Collectors.toMap(
                            country -> country.getName().toLowerCase(), // Convertir a minúsculas
                            CountryResponse::getId
                    ));
            Map<Long, Map<String, Long>> statesCache = new ConcurrentHashMap<>();
            Map<Long, Map<Long, Map<String, Long>>> citiesCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    StoreRequest storeRequest = new StoreRequest();
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Centro es requerido");
                    }
                    storeRequest.setCode(cellCode);
                    Cell denomination = row.getCell(1);
                    if(denomination == null || denomination.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Denominacion es requerido");
                    }
                    storeRequest.setDenomination(denomination.getStringCellValue());

                    String cellCountry = getCellValueAsString(row.getCell(2));
                    if (cellCountry.isEmpty()) {
                        throw new RuntimeException("Pais es requerido");
                    }
                    String cellState = getCellValueAsString(row.getCell(3));
                    if (cellState.isEmpty()) {
                        throw new RuntimeException("Estado es requerido");
                    }
                    String cellCity = getCellValueAsString(row.getCell(4));
                    if (cellCity.isEmpty()) {
                        throw new RuntimeException("Municipio es requerido");
                    }

                    Long countryId = countriesCache.get(cellCountry.toLowerCase());
                    if (countryId == null) throw new RuntimeException("Pais no encontrado: " + cellCountry);

                    synchronized (statesCache) {
                        statesCache.computeIfAbsent(countryId, id -> migrationFeign.findAllStatesByCountryId(bearerToken, id).getData().stream()
                                .collect(Collectors.toMap(
                                        state -> state.getName().toLowerCase(), // Convertir a minúsculas
                                        CountryResponse::getId
                                )));
                    }

                    Long stateId = statesCache.get(countryId).get(cellState.toLowerCase());
                    if (stateId == null) throw new RuntimeException("Estado no encontrado: " + cellState);

                    synchronized (citiesCache) {
                        citiesCache.computeIfAbsent(countryId, id -> new HashMap<>())
                                .computeIfAbsent(stateId, id -> migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId).getData().stream()
                                        .collect(Collectors.toMap(
                                                city -> city.getName().toLowerCase(), // Convertir a minúsculas
                                                CountryResponse::getId
                                        )));
                    }
                    Long cityId = citiesCache.get(countryId).get(stateId).get(cellCity.toLowerCase());
                    if (cityId == null) throw new RuntimeException("Municipio no encontrado: " + cellCity);

                    storeRequest.setCountryId(countryId);
                    storeRequest.setStateId(stateId);
                    storeRequest.setCityId(cityId);
                    storeRequest.setAddress(row.getCell(5) == null || row.getCell(5).getStringCellValue().isEmpty() ? "-" : row.getCell(5).getStringCellValue());
                    storeRequest.setZipcode(row.getCell(6) != null ? String.valueOf((int) row.getCell(6).getNumericCellValue()) : "0");
                    storeRequest.setLatitude(row.getCell(7) != null ? row.getCell(7).getNumericCellValue() : 0.0);
                    storeRequest.setLongitude(row.getCell(8) != null ? row.getCell(8).getNumericCellValue() : 0.0);
                    storeRequest.setGeorefDistance(row.getCell(9) != null ? (long) row.getCell(9).getNumericCellValue() : 0L);
                    String costCenter = getCellValueAsString(row.getCell(10));
                    Long costCenterId = null;
                    if(costCenter != null && !costCenter.isEmpty()) {
                        costCenterId = costCenterCache.computeIfAbsent(costCenter, code -> migrationFeign.findCostCenterByCode(bearerToken, code).getData().getId());
                    }
                    storeRequest.setCostCenterId(costCenterId);
                    Cell cellStatus = row.getCell(11);
                    long statusId = getStatusId(cellStatus);
                    storeRequest.setStatusId(statusId);
                    migrationFeign.createStore(bearerToken, storeRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateStoresOrgEntitiesGeographic(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursales");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> entityCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Centro es requerido");
                    }
                    Long storeId = migrationFeign.findStoreByCode(bearerToken, cellCode).getData().getId();

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;

                    String cellEmpresa = getCellValueAsString(row.getCell(12));
                    String cellRegion = getCellValueAsString(row.getCell(13));
                    String cellDivision = getCellValueAsString(row.getCell(14));
                    String cellZona = getCellValueAsString(row.getCell(15));

                    if(cellEmpresa.isEmpty()) {
                        throw new RuntimeException("Empresa es requerida");
                    }
                    String cacheKey = 1L + cellEmpresa;
                    Long empresaId;
                    if(entityCache.containsKey(cacheKey)) {
                        empresaId = entityCache.get(cacheKey);
                    } else {
                        empresaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa).getData().getId();
                        entityCache.put(cacheKey, empresaId);
                    }
                    orgEntityDetailIds.add(empresaId);
                    if (cellRegion != null && !cellRegion.isEmpty() || cellDivision != null && !cellDivision.isEmpty() || cellZona != null && !cellZona.isEmpty()) {
                        if (cellRegion != null && !cellRegion.isEmpty()) {
                            regionId = getEntityId(bearerToken, cellRegion, 2L, empresaId, "region", entityCache);
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null && !cellDivision.isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region es requerido");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division", entityCache);
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region y Division son requeridos");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Division es requerido");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona", entityCache);
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }
    public UtilResponse migrateStoresOrgEntitiesOrganizative(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursal_estructura_organizativ");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> storeCache = new ConcurrentHashMap<>();
            Map<String, Long> entityCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Centro de sucursal es requerido");
                    }
                    Long storeId = storeCache.computeIfAbsent(cellCode, code -> migrationFeign.findStoreByCode(bearerToken, code).getData().getId());

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    String cellArea = getCellValueAsString(row.getCell(1));
                    String cellSubArea = getCellValueAsString(row.getCell(2));
                    String cellDepartamento = getCellValueAsString(row.getCell(3));

                    Long subAreaId = null;
                    Long departamentoId = null;

                    if(cellArea.isEmpty()) {
                        throw new RuntimeException("Area es requerida");
                    }
                    String cacheKey = 5L + cellArea;
                    Long areaId;
                    if(entityCache.containsKey(cacheKey)) {
                        areaId = entityCache.get(cacheKey);
                    } else {
                        areaId = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea).getData().getId();
                        entityCache.put(cacheKey, areaId);
                    }
                    orgEntityDetailIds.add(areaId);
                    if (cellSubArea != null && !cellSubArea.isEmpty() || cellDepartamento != null && !cellDepartamento.isEmpty()) {
                        if (cellSubArea != null && !cellSubArea.isEmpty()) {
                            subAreaId = getEntityId(bearerToken, cellSubArea, 6L, areaId, "subarea", entityCache);
                            orgEntityDetailIds.add(subAreaId);
                        }

                        if (cellDepartamento != null && !cellDepartamento.isEmpty()) {
                            if (subAreaId == null) {
                                throw new RuntimeException("Subarea es requerida");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento", entityCache);
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_organizativ: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_organizativ: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    private long getStatusId(Cell cellStatus) {
        long statusId = 2L;
        if (cellStatus != null && !cellStatus.getStringCellValue().isEmpty()) {
            String statusValue = cellStatus.getStringCellValue().trim().toUpperCase();
            statusId = switch (statusValue) {
                case ACTIVO_STATUS -> 1L;
                case INACTIVO_STATUS -> 2L;
                default -> throw new RuntimeException("Estatus invalido: " + statusValue);
            };
        }
        return statusId;
    }

    private Long getEntityId(String bearerToken, String entityValue, Long entityType, Long parentId, String entityName, Map<String, Long> entityCache) {
        String cacheKey = entityType + "-" + (parentId != null ? parentId : "") + "-" + entityValue;
        if (entityCache.containsKey(cacheKey)) {
            return entityCache.get(cacheKey);
        }

        DefaultResponse<Page<OrgEntityResponse>> entityResponse = migrationFeign.findAllInstancesParentOrganizationEntityDetail(
                bearerToken, entityType, parentId
        );

        String name = migrationFeign.findOrgEntityDetailByName(bearerToken, entityType, entityValue).getData().getName();
        Long id = entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(name))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(entityValue).concat(" no encontrado")));

        entityCache.put(cacheKey, id); // Store in cache
        return id;
    }

    public UtilResponse migrateWorkPositions(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> workPosCatCache = new ConcurrentHashMap<>();
            Map<String, Long> storeCache = new ConcurrentHashMap<>();
            Map<String, Long> costCenterCache = new ConcurrentHashMap<>();
            Map<String, Long> organizativeStructureCache = new ConcurrentHashMap<>();
            Map<Long, DefaultResponse<StoreDetailResponse>> storeDetailsCache = new ConcurrentHashMap<>();
            Map<String, String> orgEntityCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

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
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Codigo es requerido");
                    }
                    workPositionRequest.setCode(cellCode);
                    Cell denomination = row.getCell(1);
                    if(denomination == null || denomination.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Denominacion es requerido");
                    }
                    workPositionRequest.setDenomination(denomination.getStringCellValue());
                    Cell authorizedStaff = row.getCell(2);
                    if(authorizedStaff == null || authorizedStaff.getNumericCellValue() == 0){
                        throw new RuntimeException("Plantilla autorizada es requerido");
                    }
                    workPositionRequest.setAuthorizedStaff((long)authorizedStaff.getNumericCellValue());

                    String cellWorkPosCat = getCellValueAsString(row.getCell(3));
                    if(cellWorkPosCat.isEmpty()) {
                        throw new RuntimeException("Puesto es requerido");
                    }
                    Long workPosCatId = workPosCatCache.computeIfAbsent(cellWorkPosCat, code -> migrationFeign.findWorkPosCategoryByCode(bearerToken, code).getData().getId());
                    workPositionRequest.setWorkPosCatId(workPosCatId);

                    String cellStore = getCellValueAsString(row.getCell(4));
                    if(cellStore.isEmpty()) {
                        throw new RuntimeException("Sucursal es requerido");
                    }
                    Long storeId = storeCache.computeIfAbsent(cellStore, code -> migrationFeign.findStoreByCode(bearerToken, code).getData().getId());
                    workPositionRequest.setStoreId(storeId);

                    String costCenter = getCellValueAsString(row.getCell(5));
                    Long costCenterId = null;
                    if(costCenter != null && !costCenter.isEmpty()) {
                        costCenterId = costCenterCache.computeIfAbsent(costCenter, code -> migrationFeign.findCostCenterByCode(bearerToken, code).getData().getId());
                    }
                    workPositionRequest.setCostCenterId(costCenterId);

                    Cell cellStatus = row.getCell(6);
                    long statusId = getStatusId(cellStatus);
                    workPositionRequest.setStatusId(statusId);

                    String cellArea = getCellValueAsString(row.getCell(7));
                    String cellSubarea = getCellValueAsString(row.getCell(8));
                    String cellDepartamento = getCellValueAsString(row.getCell(9));

                    if(cellArea == null || cellArea.isEmpty()) throw new RuntimeException("Area es requerida");

                    String keyStructure = cellArea + "-" + cellSubarea + "-" + cellDepartamento;
                    Long storeOrganizativeId = organizativeStructureCache.computeIfAbsent(keyStructure, key -> getStoreOrganizativeId(bearerToken, storeId, storeDetailsCache, orgEntityCache, cellArea, cellSubarea, cellDepartamento));
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
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateWorkPositionsDetails(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> workPositionCache = new ConcurrentHashMap<>();
            Map<String, Long> compCategoryCache = new ConcurrentHashMap<>();
            Map<String, Long> compTabCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode = getCellValueAsString(row.getCell(0));
                    if(cellCode.isEmpty()) {
                        throw new RuntimeException("Codigo es requerido");
                    }
                    Long workPositionId = workPositionCache.computeIfAbsent(cellCode, code ->
                            migrationFeign.findWorkPositionByCode(bearerToken, code).getData().getWorkPosition().getId()
                    );

                    String compCategory = getCellValueAsString(row.getCell(10));
                    Long compCategoryId = null;
                    if(compCategory != null && !compCategory.isEmpty()){
                        compCategoryId = compCategoryCache.computeIfAbsent(compCategory, category ->
                                migrationFeign.findCompCategoryByCode(bearerToken, category).getData().getId()
                        );
                    }

                    String compTab = getCellValueAsString(row.getCell(11));
                    Long compTabId = null;
                    if(compTab != null && !compTab.isEmpty()){
                        compTabId = compTabCache.computeIfAbsent(compTab, tab ->
                                migrationFeign.findCompTabByCode(bearerToken, tab).getData().getId()
                        );
                    }

                    String managerWorkPosition = getCellValueAsString(row.getCell(12));
                    Long managerWorkPositionId = null;
                    if(managerWorkPosition != null && !managerWorkPosition.isEmpty()){
                        managerWorkPositionId = workPositionCache.computeIfAbsent(managerWorkPosition, code ->
                                migrationFeign.findWorkPositionByCode(bearerToken, code).getData().getWorkPosition().getId()
                        );
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
                        row.getCell(0).setCellStyle(cellStyle);
                        success++;
                    }
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet cargo: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateProfiles(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empleados");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> workPositionCache = new ConcurrentHashMap<>();
            Map<String, CountryResponse> countriesCache = migrationFeign.findAll(bearerToken).getData().stream()
                    .collect(Collectors.toMap(
                            country -> country.getName().toLowerCase(), // Convertir a minúsculas para normalizar
                            country -> country
                    ));

            Map<Long, Map<String, CountryResponse>> statesCache = new ConcurrentHashMap<>();
            Map<Long, Map<Long, Map<String, CountryResponse>>> citiesCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    ProfileRequest profileRequest = new ProfileRequest();
                    List<ProfileSecValueRequest> profileSecValueRequestList = new ArrayList<>();
                    ProfileSecValueRequest informacionPersonal = new ProfileSecValueRequest();
                    informacionPersonal.setKeyword("PSPI01");
                    Map<String, Object> informacionPersonalValues = informacionPersonal.getFieldsValues();
                    String clave = getCellValueAsString(row.getCell(0));
                    if(clave.isEmpty()){
                        throw new RuntimeException("Clave MPRO es requerida");
                    }
                    informacionPersonalValues.put("Clave MPRO", clave);
                    Cell names = row.getCell(1);
                    if (names == null || names.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombres es requerido");
                    }
                    informacionPersonalValues.put("Primer Nombre", names.getStringCellValue());
                    Cell lastName = row.getCell(2);
                    if (lastName == null || lastName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Primer apellido es requerido");
                    }
                    informacionPersonalValues.put("Primer Apellido", lastName.getStringCellValue());
                    informacionPersonalValues.put("Segundo Apellido", row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "");
                    Cell gender = row.getCell(4);
                    if (gender == null || gender.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Sexo es requerido");
                    }
                    informacionPersonalValues.put("Sexo", gender.getStringCellValue());
                    DateTimeFormatter formatters = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    //LocalDate.parse(row.getCell(9).getStringCellValue(), formatters);
                    // row.getCell(9).getStringCellValue()
                    if (row.getCell(9) != null && row.getCell(9).getCellType() == CellType.NUMERIC) {
                        LocalDate hiredDate = row.getCell(9).getDateCellValue().toInstant()
                                .atZone(ZoneId.systemDefault())
                                .toLocalDate();
                        informacionPersonalValues.put("Fecha de contratación", hiredDate.format(formatters));
                    }
                    ProfileSecValueRequest informacionBiografica = new ProfileSecValueRequest();
                    informacionBiografica.setKeyword("PSBI02");
                    Map<String, Object> informacionBiograficaValues = informacionBiografica.getFieldsValues();
                    //LocalDate.parse(row.getCell(8).getStringCellValue(), formatters);
                    //row.getCell(8).getStringCellValue()
                    if (row.getCell(8) != null && row.getCell(8).getCellType() == CellType.NUMERIC) {
                        LocalDate birthDate =  row.getCell(8).getDateCellValue().toInstant()
                                .atZone(ZoneId.systemDefault())
                                .toLocalDate();
                        informacionBiograficaValues.put("Fecha de nacimiento", birthDate.format(formatters));
                    }

                    ProfileSecValueRequest datosPersonales = new ProfileSecValueRequest();
                    datosPersonales.setKeyword("PSPD03");
                    Map<String, Object> datosPersonalesValues = datosPersonales.getFieldsValues();
                    String rfc = getCellValueAsString(row.getCell(5));
                    if(rfc.isEmpty()){
                        throw new RuntimeException("RFC es requerido");
                    }
                    datosPersonalesValues.put("RFC", rfc);
                    String curp = getCellValueAsString(row.getCell(6));
                    if(curp.isEmpty()){
                        throw new RuntimeException("CURP es requerido");
                    }
                    datosPersonalesValues.put("CURP", curp);
                    String nss = getCellValueAsString(row.getCell(7));
                    if(nss.isEmpty()){
                        throw new RuntimeException("NSS es requerido");
                    }
                    datosPersonalesValues.put("NSS", nss);

                    ProfileSecValueRequest direccion = new ProfileSecValueRequest();
                    direccion.setKeyword("PSAS05");
                    Map<String, Object> direccionValues = direccion.getFieldsValues();
                    direccionValues.put("Dirección", row.getCell(12) != null ? row.getCell(12).getStringCellValue(): "");

                    String cellCountryValue = getCellValueAsString(row.getCell(13));
                    if (cellCountryValue.isEmpty()) throw new RuntimeException("Pais es requerido");
                    String cellStateValue = getCellValueAsString(row.getCell(14));
                    if (cellStateValue.isEmpty()) throw new RuntimeException("Estado es requerido");
                    String cellCityValue = getCellValueAsString(row.getCell(15));
                    if (cellCityValue.isEmpty()) throw new RuntimeException("Municipio es requerido");

                    CountryResponse paisResidencia = countriesCache.get(cellCountryValue.toLowerCase());
                    if (paisResidencia == null) throw new RuntimeException("Pais no encontrado: " + cellCountryValue);

                    synchronized (statesCache) {
                        statesCache.computeIfAbsent(paisResidencia.getId(), id ->
                                migrationFeign.findAllStatesByCountryId(bearerToken, id).getData().stream()
                                        .collect(Collectors.toMap(
                                                state -> state.getName().toLowerCase(),
                                                state -> state
                                        ))
                        );
                    }
                    CountryResponse estadoResidencia = statesCache.get(paisResidencia.getId()).get(cellStateValue.toLowerCase());
                    if (estadoResidencia == null) throw new RuntimeException("Estado no encontrado: " + cellStateValue);

                    synchronized (citiesCache) {
                        citiesCache.computeIfAbsent(paisResidencia.getId(), id -> new HashMap<>())
                                .computeIfAbsent(estadoResidencia.getId(), id ->
                                        migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, paisResidencia.getId(), estadoResidencia.getId()).getData().stream()
                                                .collect(Collectors.toMap(
                                                        city -> city.getName().toLowerCase(),
                                                        city -> city
                                                ))
                                );
                    }
                    CountryResponse ciudadResidencia = citiesCache.get(paisResidencia.getId()).get(estadoResidencia.getId()).get(cellCityValue.toLowerCase());
                    if (ciudadResidencia == null) throw new RuntimeException("Municipio no encontrado: " + cellCityValue);

                    direccionValues.put("Lugar de Residencia", Arrays.asList(paisResidencia, estadoResidencia, ciudadResidencia));

                    ProfileSecValueRequest contacto = new ProfileSecValueRequest();
                    contacto.setKeyword("PSCI06");
                    Map<String, Object> contactoValues = contacto.getFieldsValues();
                    Cell cellEmail = row.getCell(10);
                    if (cellEmail == null || cellEmail.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Email Personal es requerido");
                    }
                    contactoValues.put("Email Personal", cellEmail.getStringCellValue());
                    String cellPhone = getCellValueAsString(row.getCell(11));
                    contactoValues.put("Celular personal",  cellPhone);

                    profileSecValueRequestList.add(informacionPersonal);
                    profileSecValueRequestList.add(informacionBiografica);
                    profileSecValueRequestList.add(datosPersonales);
                    profileSecValueRequestList.add(direccion);
                    profileSecValueRequestList.add(contacto);

                    profileRequest.setSectionValues(profileSecValueRequestList);
                    String cellWorkPosition = getCellValueAsString(row.getCell(16));
                    if(cellWorkPosition.isEmpty()) {
                        throw new RuntimeException("Cargo es requerido");
                    }
                    Long workPositionId = workPositionCache.computeIfAbsent(cellWorkPosition, code ->
                            migrationFeign.findWorkPositionByCode(bearerToken, code).getData().getWorkPosition().getId()
                    );

                    profileRequest.setWorkPositionId(workPositionId);
                    migrationFeign.createProfile(bearerToken, profileRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 3);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet empleados: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateReferences(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("referencias");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Long> profilesCache = new ConcurrentHashMap<>();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    String clave = getCellValueAsString(row.getCell(0));
                    if(clave.isEmpty()){
                        throw new RuntimeException("Clave MPRO es requerida");
                    }
                    Long profileId = profilesCache.computeIfAbsent(clave, code ->
                            migrationFeign.findProfileByClaveMpro(bearerToken, code).getData().getId()
                    );

                    ProfileSecValueRequest references = new ProfileSecValueRequest();
                    references.setKeyword("PSRF16");
                    Map<String, Object> referencesValues = references.getFieldsValues();
                    Cell cellNombre = row.getCell(1);
                    String cellTelefono = getCellValueAsString(row.getCell(2));
                    if(cellNombre == null || cellNombre.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    if(cellTelefono.isEmpty()) {
                        throw new RuntimeException("Telefono es requerido");
                    }
                    referencesValues.put("Nombre", cellNombre.getStringCellValue());
                    referencesValues.put("Teléfono", cellTelefono);

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, references);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 3);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet referencias: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet referencias: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse migrateInfoBancaria(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("informacion bancaria");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    String clave = getCellValueAsString(row.getCell(0));
                    if(clave.isEmpty()){
                        throw new RuntimeException("Clave MPRO es requerida");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, clave).getData().getId();

                    ProfileSecValueRequest informacionPago = new ProfileSecValueRequest();
                    informacionPago.setKeyword("PSPM14");
                    Map<String, Object> informacionPagoValues = informacionPago.getFieldsValues();
                    Cell cellBanco = row.getCell(1);
                    String cellCuenta = getCellValueAsString(row.getCell(2));
                    String cellClabe = getCellValueAsString(row.getCell(3));
                    Cell cellTitular = row.getCell(4);
                    if(cellBanco == null || cellBanco.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Banco es requerido");
                    }
                    if(cellCuenta.isEmpty()) {
                        throw new RuntimeException("Cuenta bancaria es requerido");
                    }
                    if(cellClabe.isEmpty()) {
                        throw new RuntimeException("Clabe interbancaria es requerido");
                    }
                    if(cellTitular == null || cellTitular.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Titular de la cuenta es requerido");
                    }

                    informacionPagoValues.put("Banco", cellBanco.getStringCellValue().toUpperCase());
                    informacionPagoValues.put("Cuenta bancaria", cellCuenta);
                    informacionPagoValues.put("Clabe interbancaria", cellClabe);
                    informacionPagoValues.put("Titular de la cuenta", cellTitular.getStringCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, informacionPago);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 3);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet informacion bancaria: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet informacion bancaria: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }
    public UtilResponse migrateInfoSueldos(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sueldos");
            if(sheet == null) throw new NullException();
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            CellStyle cellStyle = this.greenCellStyle(workbook);

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    String clave = getCellValueAsString(row.getCell(0));
                    if(clave.isEmpty()){
                        throw new RuntimeException("Clave MPRO es requerida");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, clave).getData().getId();

                    ProfileSecValueRequest payrollInformation = new ProfileSecValueRequest();
                    payrollInformation.setKeyword("PSPN11");
                    Map<String, Object> payrollInformationValues = payrollInformation.getFieldsValues();
                    Cell cellSueldoMensual = row.getCell(1);
                    Cell cellSueldoDiario = row.getCell(2);
                    if(cellSueldoMensual == null || cellSueldoMensual.getNumericCellValue() == 0) {
                        throw new RuntimeException("Salario mensual es requerido");
                    }
                    if(cellSueldoDiario == null || cellSueldoMensual.getNumericCellValue() == 0) {
                        throw new RuntimeException("Sueldo diario es requerido");
                    }

                    payrollInformationValues.put("Salario mensual", cellSueldoMensual.getNumericCellValue());
                    payrollInformationValues.put("Sueldo diario", cellSueldoDiario.getNumericCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, payrollInformation);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 3);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sueldos: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet sueldos: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }
    public UtilResponse loadCompensationsCategories(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

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
                cell.setCellValue("Codigo / Denominacion / Estatus no existe");
                workbook.write(modifiedFileOutputStream);
                return new UtilResponse(modifiedFileOutputStream, success, failure);
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode1 = getCellValueAsString(row.getCell(cellCode));
                    if(cellCode1.isEmpty()){
                        throw new RuntimeException("Codigo es requerido");
                    }
                    Cell cellDenomination1 = row.getCell(cellDenomination);
                    if(cellDenomination1 == null || cellDenomination1.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Denominacion es requerido");
                    }
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
                                        if(Objects.equals(nameColumn, "Dias de aguinaldo")){
                                            fieldsValues.put(nameColumn, (long) cell.getNumericCellValue());
                                        }
                                        else {
                                            fieldsValues.put(nameColumn, cell.getNumericCellValue());
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
                    compCategories.setCode(cellCode1);
                    compCategories.setDenomination(cellDenomination1.getStringCellValue());
                    compCategories.setFieldsValues(fieldsValues);
                    compCategories.setStatusId(statusId);

                    migrationFeign.createCompensationCategories(bearerToken, compCategories);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet categorias de puesto: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet categorias de puesto: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse loadTabs(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

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
                cell.setCellValue("Nivel macropay / Posicion / Estatus no existe");
                workbook.write(modifiedFileOutputStream);
                return new UtilResponse(modifiedFileOutputStream, success, failure);
            }

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode1 = getCellValueAsString(row.getCell(cellCode));
                    if(cellCode1.isEmpty()){
                        throw new RuntimeException("Nivel macropay es requerido");
                    }
                    Cell cellDenomination1 = row.getCell(cellDenomination);
                    if(cellDenomination1 == null || cellDenomination1.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Posicion es requerido");
                    }
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
                    tabsRequest.setCode(cellCode1);
                    tabsRequest.setDenomination(cellDenomination1.getStringCellValue());
                    tabsRequest.setMinAuthorizedSalary(0L);
                    tabsRequest.setMaxAuthorizedSalary(0L);
                    tabsRequest.setStatusId(statusId);
                    tabsRequest.setFieldsValues(fieldsValues);

                    migrationFeign.createTab(bearerToken, tabsRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet tabuladores: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet tabuladores: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
    }

    public UtilResponse loadWorkPositionCategories(MultipartFile file, String bearerToken) {
        ByteArrayOutputStream modifiedFileOutputStream = new ByteArrayOutputStream();
        int success = 0;
        int failure = 0;

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
                cell.setCellValue("Codigo / Denominacion / Estatus no existe");
                workbook.write(modifiedFileOutputStream);
                return new UtilResponse(modifiedFileOutputStream, success, failure);
            }

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    String cellCode1 = getCellValueAsString(row.getCell(cellCode));
                    if(cellCode1.isEmpty()){
                        throw new RuntimeException("Codigo es requerido");
                    }
                    Cell cellDenomination1 = row.getCell(cellDenomination);
                    if(cellDenomination1 == null || cellDenomination1.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Denominacion es requerido");
                    }
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
                    workPositionCategoryRequest.setCode(cellCode1);
                    workPositionCategoryRequest.setDenomination(cellDenomination1.getStringCellValue());
                    workPositionCategoryRequest.setFieldsValues(fieldsValues);
                    workPositionCategoryRequest.setStatusId(statusId);

                    migrationFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                    success++;
                } catch (ErrorResponseException e) {
                    failure++;
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(bearerToken, sheet.getRow(i), error.getErrors(), 2);
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getError().getErrors().getFields());
                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    failure++;
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                    sheet.getRow(i).getCell(0).setCellStyle(this.redCellStyle(workbook));
                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getMessage());
                }
            }
            workbook.write(modifiedFileOutputStream);
        } catch (Exception e) {
            catchUnexpectedExceptions(e);
        }
        return new UtilResponse(modifiedFileOutputStream, success, failure);
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

    private void agregarCeldaError(Row row, String message) {
        // Agregar celda con el mensaje de error en la fila que falló
        Cell errorCell = row.createCell(row.getPhysicalNumberOfCells() + 1);
        errorCell.setCellValue(message);
    }

    private void agregarExcetionFeign(String bearerToken, Row row, ErrorDetailResponse errorDetailResponse, Integer fieldType) {
        // Agregar celda con el mensaje de error en la fila que falló
        Cell errorCell = row.createCell(row.getPhysicalNumberOfCells() + 1);

        if(errorDetailResponse == null){
            errorCell.setCellValue("Error inesperado");
            return;
        }

        String message = "";
        if (errorDetailResponse.getId() != null) {
            String name = "";
            try {
                name = migrationFeign.getNameField(bearerToken, errorDetailResponse.getId(), fieldType).getData().getName();
            } catch (Exception e1) {
                log.error("falló obteniendo el nombre del campo {}", e1.getMessage());
            }
            if (name != null && !name.isEmpty()) message = name + ": ";
        }

        for (ExceptionCodeEnum errorCodeEnum : ExceptionCodeEnum.values()) {
            if (errorCodeEnum.getCode().equals(errorDetailResponse.getCode())) {
                errorCell.setCellValue(message + errorCodeEnum.getMessage());
                return;
            }
        }

        StringBuilder errores = new StringBuilder();
        for (String f : errorDetailResponse.getFields()) {
            errores.append(f).append(" ");
        }
        errorCell.setCellValue(errores.toString());
    }

    private void catchUnexpectedExceptions(Exception e) {
        if (e instanceof FeignException e1) {
            if (e1.status() == 503) {
                throw new ServiceUnavailableException("An error occurred while communicating with the core organization service");
            }
            if (e1.status() == 401) {
                throw new UnauthorizedException("Does not contain the module permissions");
            }
            if (e1.status() == 403) {
                throw new ForbbidenException("Usuario inválido");
            }
            throw new RuntimeException(e.getMessage());
        }
        if (e instanceof NullException){
            throw new NullException("Wrong file");
        }
        if (e instanceof ErrorResponseException e1){
            throw new ErrorResponseException(e1.getError());
        }
        throw new RuntimeException(e.getMessage());
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter dataFormatter = new DataFormatter();
        return dataFormatter.formatCellValue(cell);
    }

    private Long getStoreOrganizativeId(String bearerToken, Long storeId, Map<Long, DefaultResponse<StoreDetailResponse>> storeDetailsCache, Map<String, String> orgEntityCache, String cellArea, String cellSubarea, String cellDepartamento) {
        Long storeOrganizativeId;
        DefaultResponse<StoreDetailResponse> storeDetailResponse = getStoreDetails(bearerToken, storeId, storeDetailsCache);
        //Obtener las estructuras organizativas de la sucursal cuya area sea igual a cellArea
        String area = getOrgEntityDetailName(bearerToken, 5L, cellArea, orgEntityCache);
        List<OrgEntDetailResponse> areasFiltradas = storeDetailResponse.getData().getStructuresByType().stream()
                .flatMap(structureType -> structureType.getDetails().stream())
                .filter(detail -> detail.getStructures().stream().anyMatch(structure -> area.equalsIgnoreCase(structure.getName()) && structure.getOrgEntity().getId() == 5L))
                .toList();
        //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa area
        if (areasFiltradas.isEmpty()) throw new RuntimeException("Area ".concat(cellArea).concat(" no encontrada. Debe coincidir con la estructura de la sucursal."));

        if (cellSubarea != null && !cellSubarea.isEmpty()) {
            //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas tienen el subarea
            String subArea = getOrgEntityDetailName(bearerToken, 6L, cellSubarea, orgEntityCache);
            List<OrgEntDetailResponse> areasFiltradasConSubarea = areasFiltradas.stream()
                    .filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> subArea.equalsIgnoreCase(child.getName()) && child.getOrgEntity().getId() == 6L)))
                    .toList();
            //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa subarea
            if (areasFiltradasConSubarea.isEmpty()) throw new RuntimeException("Subarea ".concat(cellSubarea).concat(" no encontrada. Debe coincidir con la estructura de la sucursal."));

            if (cellDepartamento != null && !cellDepartamento.isEmpty()) {
                //Una vez encontradas las estructuras organizativas que tienen ese area-subarea, buscar cual de ellas tienen el departamento
                String departamento = getOrgEntityDetailName(bearerToken, 7L, cellDepartamento, orgEntityCache);
                Optional<OrgEntDetailResponse> areaConSubareaYDepartamento = areasFiltradasConSubarea.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> subArea.equalsIgnoreCase(child.getName()) && child.getOrgEntity().getId() == 6L && child.getChildren().stream().anyMatch(child2 -> child2.getName().equalsIgnoreCase(departamento) && child2.getOrgEntity().getId() == 7L))))
                        .findFirst();

                //Si el optional es vacio es porque ningun area-subarea tiene ese departamento
                if (areaConSubareaYDepartamento.isEmpty()) throw new RuntimeException("Departamento ".concat(cellDepartamento).concat(" no encontrado. Debe coincidir con la estructura de la sucursal"));
                storeOrganizativeId = areaConSubareaYDepartamento.get().getId();
            }
            else {
                //Una vez encontradas las estructuras organizativas que tienen ese area-subarea, buscar cual de ellas no tiene departamento
                Optional<OrgEntDetailResponse> areaConSubareaSinDepartamento = areasFiltradasConSubarea.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren().stream().anyMatch(child -> child.getChildren() == null || child.getChildren().isEmpty() || child.getChildren().get(0) == null)))
                        .findFirst();
                //Si el optional es vacio es porque todas las area-subarea tienen un departamento y se necesita que en el excel se envíe el departamento para buscarlo
                if(areaConSubareaSinDepartamento.isEmpty()) throw new RuntimeException("Departamento es requerido");
                storeOrganizativeId = areaConSubareaSinDepartamento.get().getId();
            }
        }
        else  {
            //Si el excel tiene un departamento y no tiene un subarea, entonces está mal la estructura, falta el subarea
            if(cellDepartamento != null && !cellDepartamento.isEmpty()) throw new RuntimeException("Subarea es requerida");

            //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas no tiene subarea
            Optional<OrgEntDetailResponse> areaSinSubarea = areasFiltradas.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() == null || structure.getChildren().isEmpty() || structure.getChildren().get(0) == null))
                    .findFirst();
            //Si el optional es vacio es porque todas las areas tienen un subarea y se necesita que en el excel se envíe el subarea para buscarlo
            if(areaSinSubarea.isEmpty()) throw new RuntimeException("Subarea es requerida");
            storeOrganizativeId = areaSinSubarea.get().getId();
        }
        return storeOrganizativeId;
    }

    public DefaultResponse<StoreDetailResponse> getStoreDetails(String token, Long storeId, Map<Long, DefaultResponse<StoreDetailResponse>> storeDetailsCache) {
        if (!storeDetailsCache.containsKey(storeId)) {
            storeDetailsCache.put(storeId, migrationFeign.findAllStoresDetails(token, storeId));
        }
        return storeDetailsCache.get(storeId);
    }

    private String getOrgEntityDetailName(String token, Long entityId, String name, Map<String, String> orgEntityCache) {
        String cacheKey = entityId + "-" + name;
        if (!orgEntityCache.containsKey(cacheKey)) {
            orgEntityCache.put(cacheKey, migrationFeign.findOrgEntityDetailByName(token, entityId, name).getData().getName());
        }
        return orgEntityCache.get(cacheKey);
    }
}
