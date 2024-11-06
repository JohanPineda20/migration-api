package com.nelumbo.migration.service;

import com.nelumbo.migration.exceptions.ErrorResponseException;
import com.nelumbo.migration.exceptions.NullCellException;
import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import com.nelumbo.migration.feign.dto.responses.error.ErrorDetailResponse;
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
    private static final String MODIFIED = "modified_";
    private static final String SHEET = "Estamos con la hoja: ";
    private static final String COUNTROWS = "La cantidad de filas es: ";

    private final MigrationFeign migrationFeign;

    public List<ErrorResponse> migrateEmpresa(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) -> {
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 1L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateRegion(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellEmpresa = row.getCell(1);
                    if(cellEmpresa == null || cellEmpresa.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Empresa es requerida");
                    }
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 1L, cellEmpresa.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 2L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateDivision(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellRegion = row.getCell(1);
                    if(cellRegion == null || cellRegion.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Region es requerido");
                    }
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 2L, cellRegion.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 3L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateZona(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellDivision = row.getCell(1);
                    if(cellDivision == null || cellDivision.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Division es requerido");
                    }
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 3L, cellDivision.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 4L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateArea(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 5L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateSubarea(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellArea = row.getCell(1);
                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Area es requerida");
                    }
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 6L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateDepartamento(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellName == null || cellName.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    orgEntityDetailRequest.setName(cellName.getStringCellValue());
                    Cell cellSubarea = row.getCell(1);
                    if(cellSubarea == null || cellSubarea.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Subarea es requerida");
                    }
                    orgEntityDetailRequest.setParentId(migrationFeign.findOrgEntityDetailByName(bearerToken, 6L, cellSubarea.getStringCellValue()).getData().getId());
                    fieldValues.forEach((name, position) ->{
                        Cell cell = row.getCell(position);
                        if (cell != null) {
                            orgEntityDetailRequest.getFieldValues().put(name, cell.getStringCellValue());
                        }
                    });

                    migrationFeign.createOrgEntityDetail(bearerToken, orgEntityDetailRequest, 7L);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateCostCenters(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(cellCode == null || cellCode.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Codigo es requerido");
                    }
                    costCenterRequest.setCode(cellCode.getStringCellValue());
                    Cell cellDenomination = row.getCell(1);
                    if(cellDenomination == null || cellDenomination.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Denominacion es requerida");
                    }
                    costCenterRequest.setDenomination(cellDenomination.getStringCellValue());

                    Cell cellCountry = row.getCell(2);
                    if(cellCountry == null || cellCountry.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Pais es requerido");
                    }
                    Cell cellState = row.getCell(3);
                    if(cellState == null || cellState.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Estado es requerido");
                    }
                    Cell cellCity = row.getCell(4);
                    if(cellCity == null || cellCity.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Municipio es requerido");
                    }

                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(cellCountry.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("Pais ".concat(cellCountry.getStringCellValue().concat(" no encontrado"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken,countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(cellState.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("Estado ".concat(cellState.getStringCellValue().concat(" no encontrado"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(cellCity.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("Municipio ".concat(cellCity.getStringCellValue().concat(" no encontrado"))));

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
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateCostCentersOrgEntitiesGeographic(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("ceco_estructura_geografica");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    Cell cellCode = row.getCell(0);
                    if(cellCode == null || cellCode.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Codigo del Centro de Costos es requerido");
                    }
                    Long costCenterId = migrationFeign.findCostCenterByCode(bearerToken, cellCode.getStringCellValue()).getData().getId();

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
                        throw new RuntimeException("Empresa es requerida");
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
                                throw new RuntimeException("Region es requerido");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region y Division son requeridos");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Division es requerido");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona");
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateCostCentersOrgEntitiesOrganizative(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                        throw new RuntimeException("Area es requerida");
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
                                throw new RuntimeException("Subarea es requerida");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento");
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createCostCenterDetails(bearerToken, costCenterDetailRequest, costCenterId);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateStores(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sucursales");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    StoreRequest storeRequest = new StoreRequest();
                    Cell code = row.getCell(0);
                    if(code == null || code.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Centro es requerido");
                    }
                    storeRequest.setCode(code.getStringCellValue());
                    Cell denomination = row.getCell(1);
                    if(denomination == null || denomination.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Denominacion es requerido");
                    }
                    storeRequest.setDenomination(denomination.getStringCellValue());

                    Cell cellCountry = row.getCell(2);
                    if(cellCountry == null || cellCountry.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Pais es requerido");
                    }
                    Cell cellState = row.getCell(3);
                    if(cellState == null || cellState.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Estado es requerido");
                    }
                    Cell cellCity = row.getCell(4);
                    if(cellCity == null || cellCity.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Municipio es requerido");
                    }

                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(cellCountry.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("country ".concat(cellCountry.getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken, countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(cellState.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("state ".concat(cellState.getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(cellCity.getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("city ".concat(cellCity.getStringCellValue().concat(" not found"))));

                    storeRequest.setCountryId(countryId);
                    storeRequest.setStateId(stateId);
                    storeRequest.setCityId(cityId);
                    storeRequest.setAddress(row.getCell(5) == null || row.getCell(5).getStringCellValue().isEmpty() ? "-" : row.getCell(5).getStringCellValue());
                    storeRequest.setZipcode(row.getCell(6) != null ? String.valueOf((int) row.getCell(6).getNumericCellValue()) : "00000");
                    storeRequest.setLatitude(row.getCell(7) != null ? row.getCell(7).getNumericCellValue() : 0.0);
                    storeRequest.setLongitude(row.getCell(8) != null ? row.getCell(8).getNumericCellValue() : 0.0);
                    storeRequest.setGeorefDistance(row.getCell(9) != null ? (long) row.getCell(9).getNumericCellValue() : 0L);
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
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateStoresOrgEntitiesGeographic(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursales");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Cell code = row.getCell(0);
                    if(code == null || code.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Centro es requerido");
                    }
                    Long storeId = migrationFeign.findStoreByCode(bearerToken, code.getStringCellValue()).getData().getId();

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;

                    Cell cellEmpresa = row.getCell(12);
                    Cell cellRegion = row.getCell(13);
                    Cell cellDivision = row.getCell(14);
                    Cell cellZona = row.getCell(15);

                    if(cellEmpresa == null || cellEmpresa.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Empresa es requerida");
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
                                throw new RuntimeException("Region es requerido");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null && !cellZona.getStringCellValue().isEmpty()) {
                            if (regionId == null) {
                                throw new RuntimeException("Region y Division son requeridos");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Division es requerido");
                            }
                            zonaId = getEntityId(bearerToken, cellZona, 4L, divisionId, "zona");
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateStoresOrgEntitiesOrganizative(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursal_estructura_organizativ");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    Cell code = row.getCell(0);
                    if(code == null || code.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Centro de Sucursal es requerido");
                    }
                    Long storeId = migrationFeign.findStoreByCode(bearerToken, code.getStringCellValue()).getData().getId();

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Cell cellArea = row.getCell(1);
                    Cell cellSubArea = row.getCell(2);
                    Cell cellDepartamento = row.getCell(3);

                    Long subAreaId = null;
                    Long departamentoId = null;

                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Area es requerida");
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
                                throw new RuntimeException("Subarea es requerida");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 7L, subAreaId, "departamento");
                            orgEntityDetailIds.add(departamentoId);
                        }
                    }
                    migrationFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
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

    private Long getEntityId(String bearerToken, Cell cell, Long entityType, Long parentId, String entityName) {
        DefaultResponse<Page<OrgEntityResponse>> entityResponse = migrationFeign.findAllInstancesParentOrganizationEntityDetail(
                bearerToken, entityType, parentId
        );

        String name = migrationFeign.findOrgEntityDetailByName(bearerToken, entityType, cell.getStringCellValue()).getData().getName();

        return entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(name))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(cell.getStringCellValue()).concat(" no encontrado")));
    }

    public List<ErrorResponse> migrateWorkPositions(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    if(code == null || code.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Codigo es requerido");
                    }
                    workPositionRequest.setCode(code.getStringCellValue());
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

                    Cell cellWorkPosCat = row.getCell(3);
                    if(cellWorkPosCat == null || cellWorkPosCat.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Puesto es requerido");
                    }
                    Long workPosCatId = migrationFeign.findWorkPosCategoryByCode(bearerToken, cellWorkPosCat.getStringCellValue()).getData().getId();
                    workPositionRequest.setWorkPosCatId(workPosCatId);

                    Cell cellStore = row.getCell(4);
                    if(cellStore == null || cellStore.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Sucursal es requerida");
                    }
                    Long storeId = migrationFeign.findStoreByCode(bearerToken, cellStore.getStringCellValue()).getData().getId();
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

                    if(cellArea == null || cellArea.getStringCellValue().isEmpty()) throw new RuntimeException("Area es requerida");

                    DefaultResponse<StoreDetailResponse> storeDetailResponse = migrationFeign.findAllStoresDetails(bearerToken, storeId);
                    //Obtener las estructuras organizativas de la sucursal cuya area sea igual a cellArea
                    String area = migrationFeign.findOrgEntityDetailByName(bearerToken, 5L, cellArea.getStringCellValue()).getData().getName();
                    List<OrgEntDetailResponse> areasFiltradas = storeDetailResponse.getData().getStructuresByType().stream()
                            .flatMap(structureType -> structureType.getDetails().stream())
                            .filter(detail -> detail.getStructures().stream().anyMatch(structure -> area.equalsIgnoreCase(structure.getName()) && structure.getOrgEntity().getId() == 5L))
                            .toList();
                    //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa area
                    if (areasFiltradas.isEmpty()) throw new RuntimeException("Area ".concat(cellArea.getStringCellValue()).concat(" no encontrada. Debe coincidir con la estructura de la sucursal."));

                    if (cellSubarea != null && !cellSubarea.getStringCellValue().isEmpty()) {
                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas tienen el subarea
                        String subArea = migrationFeign.findOrgEntityDetailByName(bearerToken, 6L, cellSubarea.getStringCellValue()).getData().getName();
                        List<OrgEntDetailResponse> areasFiltradasConSubarea = areasFiltradas.stream()
                                .filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> subArea.equalsIgnoreCase(child.getName()) && child.getOrgEntity().getId() == 6L)))
                                .toList();
                        //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa subarea
                        if (areasFiltradasConSubarea.isEmpty()) throw new RuntimeException("Subarea ".concat(cellSubarea.getStringCellValue()).concat(" no encontrada. Debe coincidir con la estructura de la sucursal."));

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
                            if(areaConSubareaSinDepartamento.isEmpty()) throw new RuntimeException("Departamento es requerido");
                            storeOrganizativeId = areaConSubareaSinDepartamento.get().getId();
                        }
                    }
                    else  {
                        //Si el excel tiene un departamento y no tiene un subarea, entonces está mal la estructura, falta el subarea
                        if(cellDepartamento != null && !cellDepartamento.getStringCellValue().isEmpty()) throw new RuntimeException("Subarea es requerida");

                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas no tiene subarea
                        Optional<OrgEntDetailResponse> areaSinSubarea = areasFiltradas.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() == null || structure.getChildren().isEmpty() || structure.getChildren().get(0) == null))
                                .findFirst();
                        //Si el optional es vacio es porque todas las areas tienen un subarea y se necesita que en el excel se envíe el subarea para buscarlo
                        if(areaSinSubarea.isEmpty()) throw new RuntimeException("Subarea es requerida");
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
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateWorkPositionsDetails(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    Cell code = row.getCell(0);
                    if(code == null || code.getStringCellValue().isEmpty()){
                        throw new RuntimeException("Codigo es requerido");
                    }
                    Long workPositionId = migrationFeign.findWorkPositionByCode(bearerToken, code.getStringCellValue()).getData().getWorkPosition().getId();
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
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateProfiles(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empleados");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    ProfileRequest profileRequest = new ProfileRequest();
                    List<ProfileSecValueRequest> profileSecValueRequestList = new ArrayList<>();
                    ProfileSecValueRequest informacionPersonal = new ProfileSecValueRequest();
                    informacionPersonal.setKeyword("PSPI01");
                    Map<String, Object> informacionPersonalValues = informacionPersonal.getFieldsValues();
                    Cell clave = row.getCell(0);
                    if (clave == null || clave.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO es requerido");
                    }
                    informacionPersonalValues.put("Clave MPRO", clave.getStringCellValue());
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
                    if (row.getCell(9) != null) {
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
                    if (row.getCell(8) != null) {
                        LocalDate birthDate =  row.getCell(8).getDateCellValue().toInstant()
                                .atZone(ZoneId.systemDefault())
                                .toLocalDate();
                        informacionBiograficaValues.put("Fecha de nacimiento", birthDate.format(formatters));
                    }

                    ProfileSecValueRequest datosPersonales = new ProfileSecValueRequest();
                    datosPersonales.setKeyword("PSPD03");
                    Map<String, Object> datosPersonalesValues = datosPersonales.getFieldsValues();
                    datosPersonalesValues.put("RFC", row.getCell(5) != null ? row.getCell(5).getStringCellValue() : "");
                    datosPersonalesValues.put("CURP", row.getCell(6) != null ? row.getCell(6).getStringCellValue(): "");
                    datosPersonalesValues.put("NSS", row.getCell(7) != null ? row.getCell(7).getStringCellValue(): "");

                    ProfileSecValueRequest direccion = new ProfileSecValueRequest();
                    direccion.setKeyword("PSAS05");
                    Map<String, Object> direccionValues = direccion.getFieldsValues();
                    direccionValues.put("Dirección", row.getCell(12) != null ? row.getCell(12).getStringCellValue(): "");
                    Cell cellCountry = row.getCell(13);
                    if (cellCountry == null || cellCountry.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Pais es requerido");
                    }
                    Cell cellState = row.getCell(14);
                    if (cellState == null || cellState.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Estado es requerido");
                    }
                    Cell cellCity = row.getCell(15);
                    if (cellCity == null || cellCity.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Municipio es requerido");
                    }
                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    CountryResponse paisResidencia = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(cellCountry.getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("country ".concat(cellCountry.getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken, paisResidencia.getId());
                    CountryResponse estadoResidencia = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(cellState.getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("state ".concat(cellState.getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, paisResidencia.getId(), estadoResidencia.getId());
                    CountryResponse ciudadResidencia = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(cellCity.getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("city ".concat(cellCity.getStringCellValue().concat(" not found"))));
                    direccionValues.put("Lugar de Residencia", Arrays.asList(paisResidencia, estadoResidencia, ciudadResidencia));

                    ProfileSecValueRequest contacto = new ProfileSecValueRequest();
                    contacto.setKeyword("PSCI06");
                    Map<String, Object> contactoValues = contacto.getFieldsValues();
                    Cell cellEmail = row.getCell(10);
                    if (cellEmail == null || cellEmail.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Email Personal es requerido");
                    }
                    contactoValues.put("Email Personal", cellEmail.getStringCellValue());
                    contactoValues.put("Celular personal", row.getCell(11) == null ? "" : row.getCell(11).getStringCellValue());

                    profileSecValueRequestList.add(informacionPersonal);
                    profileSecValueRequestList.add(informacionBiografica);
                    profileSecValueRequestList.add(datosPersonales);
                    profileSecValueRequestList.add(direccion);
                    profileSecValueRequestList.add(contacto);

                    profileRequest.setSectionValues(profileSecValueRequestList);
                    Cell cellWorkPosition = row.getCell(16);
                    if (cellWorkPosition == null || cellWorkPosition.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Cargo es requerido");
                    }
                    Long workPositionId = migrationFeign.findWorkPositionByCode(bearerToken, cellWorkPosition.getStringCellValue()).getData().getWorkPosition().getId();

                    profileRequest.setWorkPositionId(workPositionId);
                    migrationFeign.createProfile(bearerToken, profileRequest);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> migrateProfilesGroups(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateReferences(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("referencias");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO es requerido");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, cellClaveMPRO.getStringCellValue()).getData().getId();

                    ProfileSecValueRequest references = new ProfileSecValueRequest();
                    references.setKeyword("PSRF16");
                    Map<String, Object> referencesValues = references.getFieldsValues();
                    Cell cellNombre = row.getCell(1);
                    Cell cellTelefono = row.getCell(2);
                    if(cellNombre == null || cellNombre.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Nombre es requerido");
                    }
                    if(cellTelefono == null || cellTelefono.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Telefono es requerido");
                    }
                    referencesValues.put("Nombre", cellNombre.getStringCellValue());
                    referencesValues.put("Teléfono", cellTelefono.getStringCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, references);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateInfoBancaria(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("informacion bancaria");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO es requerido");
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
                        throw new RuntimeException("Banco es requerido");
                    }
                    if(cellCuenta == null || cellCuenta.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Cuenta bancaria es requerido");
                    }
                    if(cellClabe == null || cellClabe.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clabe interbancaria es requerido");
                    }
                    if(cellTitular == null || cellTitular.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Titular de la cuenta es requerido");
                    }

                    informacionPagoValues.put("Banco", cellBanco.getStringCellValue().toUpperCase());
                    informacionPagoValues.put("Cuenta bancaria", cellCuenta.getStringCellValue());
                    informacionPagoValues.put("Clabe interbancaria", cellClabe.getStringCellValue());
                    informacionPagoValues.put("Titular de la cuenta", cellTitular.getStringCellValue());

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, informacionPago);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> migrateInfoSueldos(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sueldos");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    Cell cellClaveMPRO = row.getCell(0);
                    if (cellClaveMPRO == null || cellClaveMPRO.getStringCellValue().isEmpty()) {
                        throw new RuntimeException("Clave MPRO es requerido");
                    }
                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, cellClaveMPRO.getStringCellValue()).getData().getId();

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
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    public List<ErrorResponse> loadCompensationsCategories(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    Cell cellCode1 = row.getCell(cellCode);
                    if(cellCode1 == null || cellCode1.getStringCellValue().isEmpty()){
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
                    compCategories.setCode(cellCode1.getStringCellValue());
                    compCategories.setDenomination(cellDenomination1.getStringCellValue());
                    compCategories.setFieldsValues(fieldsValues);
                    compCategories.setStatusId(statusId);

                    migrationFeign.createCompensationCategories(bearerToken, compCategories);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> loadTabs(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    Cell cellCode1 = row.getCell(cellCode);
                    if(cellCode1 == null || cellCode1.getStringCellValue().isEmpty()){
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
                    tabsRequest.setCode(cellCode1.getStringCellValue());
                    tabsRequest.setDenomination(cellDenomination1.getStringCellValue());
                    tabsRequest.setMinAuthorizedSalary(0L);
                    tabsRequest.setMaxAuthorizedSalary(0L);
                    tabsRequest.setStatusId(statusId);
                    tabsRequest.setFieldsValues(fieldsValues);

                    migrationFeign.createTab(bearerToken, tabsRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja tabuladores");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja tabuladores: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet tabuladores: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }

    public List<ErrorResponse> loadWorkPositionCategories(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                    Cell cellCode1 = row.getCell(cellCode);
                    if(cellCode1 == null || cellCode1.getStringCellValue().isEmpty()){
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
                    workPositionCategoryRequest.setCode(cellCode1.getStringCellValue());
                    workPositionCategoryRequest.setDenomination(cellDenomination1.getStringCellValue());
                    workPositionCategoryRequest.setFieldsValues(fieldsValues);
                    workPositionCategoryRequest.setStatusId(statusId);

                    migrationFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja puestos");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja puestos: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet puestos: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
    }
    
    public List<ErrorResponse> loadGroups(MultipartFile file, String bearerToken) {
        List<ErrorResponse> errors = new ArrayList<>();
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
                } catch (ErrorResponseException e) {
                    ErrorResponse error = e.getError();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones");
                    errors.add(error);

                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    ErrorResponse error = new ErrorResponse();
                    error.setMessage("Error procesando fila " + (i + 1) + " en la hoja regiones: " + e.getMessage());
                    ErrorDetailResponse errorDetail = new ErrorDetailResponse();
                    errorDetail.setCode("C03");
                    errorDetail.setDescription("Validation Exception");
                    errorDetail.setFields(Collections.singletonList(e.getMessage()));
                    error.setErrors(errorDetail);
                    errors.add(error);
                    log.error("Error processing row " + (i + 1) + " in sheet regiones: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            ErrorResponse error = new ErrorResponse();
            error.setMessage("Error procesando archivo " + e.getMessage());
            errors.add(error);
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return errors;
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
