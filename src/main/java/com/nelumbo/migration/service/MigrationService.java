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
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class MigrationService {
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

    public void migrateCostCenters(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("centro de costos");
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String, Integer> fieldValues = new HashMap<>();
            Row rowEncabezados = sheet.getRow(0);
            for(int i = 5; i < rowEncabezados.getPhysicalNumberOfCells(); i++) {
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
                    costCenterRequest.setStatusId(1L);
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
                    storeRequest.setStatusId(1L);
                    storeRequest.setAddress(row.getCell(5).getStringCellValue());
                    storeRequest.setZipcode("" + (int) row.getCell(6).getNumericCellValue());
                    storeRequest.setLatitude(row.getCell(7).getNumericCellValue());
                    storeRequest.setLongitude(row.getCell(8).getNumericCellValue());
                    storeRequest.setGeorefDistance((long) row.getCell(9).getNumericCellValue());
                    String costCenter = row.getCell(10) != null ? row.getCell(10).getStringCellValue() : null;
                    Long costCenterId = null;
                    if(costCenter != null) {
                        costCenterId = migrationFeign.findCostCenterByCode(bearerToken, costCenter).getData().getId();
                    }
                    storeRequest.setCostCenterId(costCenterId);
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
            Sheet sheet = workbook.getSheet("sucursal_estructura_geografica");
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

                    Cell cellRegion = row.getCell(2);
                    Cell cellDivision = row.getCell(3);
                    Cell cellZona = row.getCell(4);

                    orgEntityDetailIds.add(1L);
                    if (cellRegion != null || cellDivision != null || cellZona != null) {
                        if (cellRegion != null) {
                            regionId = getEntityId(bearerToken, cellRegion, 2L, 1L, "region");
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region");
                            }
                            divisionId = getEntityId(bearerToken, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null) {
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
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_geografica: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_estructura_geografica: " + e.getMessage());
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

                    Long areaId = null;
                    Long departamentoId = null;

                    Cell cellArea = row.getCell(1);
                    Cell cellDepartamento = row.getCell(2);


                        if (cellArea != null) {
                            areaId = getEntityId(bearerToken, cellArea, 5L, 0L, "area");
                            orgEntityDetailIds.add(areaId);
                        }

                        if (cellDepartamento != null) {
                            if (areaId == null) {
                                throw new RuntimeException("Invalid organizative structure: missing area");
                            }
                            departamentoId = getEntityId(bearerToken, cellDepartamento, 6L, areaId, "departamento");
                            orgEntityDetailIds.add(departamentoId);
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
    private Long getEntityId(String bearerToken, Cell cell, Long entityType, Long parentId, String entityName) {
        DefaultResponse<Page<OrgEntityResponse>> entityResponse = migrationFeign.findAllInstancesParentOrganizationEntityDetail(
                bearerToken, entityType, parentId
        );

        return entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(cell.getStringCellValue()))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(cell.getStringCellValue()).concat(" not found")));
    }

    public void migrateStoreWorkPeriods(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("sucursal_jornadas");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    StoreWorkPeriodRequest storeWorkPeriodRequest = new StoreWorkPeriodRequest();
                    Long storeId = migrationFeign.findStoreByCode(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();
                    Long workPeriodId = migrationFeign.findWorkPeriodByNameAndWorkPeriodType(bearerToken, row.getCell(1).getStringCellValue(), "Fixed Scheduled/Regular shift").getData().getId();
                    storeWorkPeriodRequest.setWorkPeriodId(workPeriodId);
                    migrationFeign.createStoreWorkPeriods(bearerToken, storeWorkPeriodRequest, storeId);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_jornadas: " + e.getError().getErrors().getFields());

                    if (e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_jornadas: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
    public void migrateWorkPositions(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("cargo");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    WorkPositionRequest workPositionRequest = new WorkPositionRequest();
                    Cell code = row.getCell(0);
                    workPositionRequest.setCode(code.getStringCellValue());
                    workPositionRequest.setDenomination(row.getCell(1).getStringCellValue());
                    workPositionRequest.setAuthorizedStaff((long)row.getCell(2).getNumericCellValue());
                    workPositionRequest.setStatusId(1L);

                    Long workPosCatId = migrationFeign.findWorkPosCategoryByCode(bearerToken, row.getCell(3).getStringCellValue()).getData().getId();
                    workPositionRequest.setWorkPosCatId(workPosCatId);

                    Long storeId = migrationFeign.findStoreByCode(bearerToken, row.getCell(4).getStringCellValue()).getData().getId();
                    workPositionRequest.setStoreId(storeId);

                    Long storeOrganizativeId = null;
                    Cell cellArea = row.getCell(5);
                    if(cellArea == null) throw new RuntimeException("Area is required");
                    DefaultResponse<StoreDetailResponse> storeDetailResponse = migrationFeign.findAllStoresDetails(bearerToken, storeId);
                    //Obtener las estructuras organizativas de la sucursal cuya area sea igual a cellArea
                    List<OrgEntDetailResponse> areasFiltradas = storeDetailResponse.getData().getStructuresByType().stream()
                            .flatMap(structureType -> structureType.getDetails().stream())
                            .filter(detail -> detail.getStructures().stream().anyMatch(structure -> cellArea.getStringCellValue().equalsIgnoreCase(structure.getName())))
                            .toList();
                    //Si la lista es vacia es porque ninguna de las estructuras organizativas de la sucursal tiene esa area
                    if (areasFiltradas.isEmpty()) throw new RuntimeException("Area ".concat(cellArea.getStringCellValue()).concat(" not found"));

                    Cell cellDepartamento = row.getCell(6);
                    if (cellDepartamento != null) {
                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas tienen el departamento
                        Optional<OrgEntDetailResponse> areaConDepartamento = areasFiltradas.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() != null && !structure.getChildren().isEmpty() && structure.getChildren().get(0) != null && structure.getChildren().stream().anyMatch(child -> cellDepartamento.getStringCellValue().equalsIgnoreCase(child.getName()))))
                                .findFirst();
                        //Si el optional es vacio es porque ningun area tiene ese departamento
                        if (areaConDepartamento.isEmpty()) throw new RuntimeException("Departamento ".concat(cellDepartamento.getStringCellValue()).concat(" not found"));
                        storeOrganizativeId = areaConDepartamento.get().getId();
                    }
                    else  {
                        //Una vez encontradas las estructuras organizativas que tienen ese area, buscar cual de ellas no tiene departamento
                        Optional<OrgEntDetailResponse> areaSinDepartamento = areasFiltradas.stream().filter(detail -> detail.getStructures().stream().anyMatch(structure -> structure.getChildren() == null || structure.getChildren().isEmpty() || structure.getChildren().get(0) == null))
                                .findFirst();
                        //Si el optional es vacio es porque todas las areas tienen un departamento y se necesita que en el excel se envíe el departamento para buscarlo
                        if(areaSinDepartamento.isEmpty()) throw new RuntimeException("A departamento is required");
                        storeOrganizativeId = areaSinDepartamento.get().getId();
                    }
                    workPositionRequest.setStoreOrganizativeId(storeOrganizativeId);
                    String costCenter = row.getCell(7) != null ? row.getCell(7).getStringCellValue() : null;
                    Long costCenterId = null;
                    if(costCenter != null) {
                        costCenterId = migrationFeign.findCostCenterByCode(bearerToken, costCenter).getData().getId();
                    }
                    workPositionRequest.setCostCenterId(costCenterId);
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
                    String compCategory = row.getCell(8) != null ? row.getCell(8).getStringCellValue() : null;
                    Long compCategoryId = null;
                    if(compCategory != null){
                        compCategoryId = migrationFeign.findCompCategoryByCode(bearerToken, compCategory).getData().getId();
                    }
                    String compTab = row.getCell(9) != null ? row.getCell(9).getStringCellValue() : null;
                    Long compTabId = null;
                    if(compTab != null){
                        compTabId = migrationFeign.findCompTabByCode(bearerToken, compTab).getData().getId();
                    }
                    Long minSalary = row.getCell(10) != null ? (long)row.getCell(10).getNumericCellValue() : null;
                    WorkPositionUpdateRequest wPUReq = WorkPositionUpdateRequest.builder()
                            .compCategoryId(compCategoryId)
                            .compTabId(compTabId)
                            .minSalary(minSalary)
                            .build();
                    if(!(compCategoryId == null && compTabId == null && minSalary == null)){
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
                    informacionPersonalValues.put("Segundo Nombre", row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "");
                    informacionPersonalValues.put("Primer Apellido", row.getCell(3).getStringCellValue());
                    informacionPersonalValues.put("Segundo Apellido", row.getCell(4) != null ? row.getCell(4).getStringCellValue() : "");
                    informacionPersonalValues.put("Sexo", row.getCell(5).getStringCellValue());
                    informacionPersonalValues.put("Grupo sanguíneo", row.getCell(12).getStringCellValue());
                    informacionPersonalValues.put("Estado civil", row.getCell(14).getStringCellValue());
                    DateTimeFormatter formatters = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    LocalDate.parse(row.getCell(15).getStringCellValue(), formatters);
                    informacionPersonalValues.put("Fecha de ingreso", row.getCell(15).getStringCellValue());
                    informacionPersonalValues.put("Clave MPRO", clave.getStringCellValue());

                    ProfileSecValueRequest informacionBiografica = new ProfileSecValueRequest();
                    informacionBiografica.setKeyword("PSBI02");
                    Map<String, Object> informacionBiograficaValues = informacionBiografica.getFieldsValues();
                    LocalDate.parse(row.getCell(13).getStringCellValue(), formatters);
                    informacionBiograficaValues.put("Fecha de nacimiento", row.getCell(13).getStringCellValue());

                    ProfileSecValueRequest datosPersonales = new ProfileSecValueRequest();
                    datosPersonales.setKeyword("PSPD03");
                    Map<String, Object> datosPersonalesValues = datosPersonales.getFieldsValues();
                    datosPersonalesValues.put("RFC", row.getCell(6).getStringCellValue());
                    datosPersonalesValues.put("CURP", row.getCell(7).getStringCellValue());
                    datosPersonalesValues.put("NSS", row.getCell(8).getStringCellValue());
                    datosPersonalesValues.put("INFONAVIT", row.getCell(9) != null ? row.getCell(9).getStringCellValue() : "");
                    datosPersonalesValues.put("FONACOT", row.getCell(10) != null ? row.getCell(10).getStringCellValue(): "");
                    datosPersonalesValues.put("ST", row.getCell(11) != null ? row.getCell(11).getStringCellValue(): ""); //no se sabe que es, tampoco esta en base de datos

                    ProfileSecValueRequest direccion = new ProfileSecValueRequest();
                    direccion.setKeyword("PSAS05");
                    Map<String, Object> direccionValues = direccion.getFieldsValues();
                    direccionValues.put("Dirección", row.getCell(19).getStringCellValue());
                    direccionValues.put("Transporte", (row.getCell(24) == null) ? "" : row.getCell(24).getStringCellValue());
                    DefaultResponse<List<CountryResponse>> countryResponse = migrationFeign.findAll(bearerToken);
                    CountryResponse paisResidencia = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(20).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(20).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = migrationFeign.findAllStatesByCountryId(bearerToken, paisResidencia.getId());
                    CountryResponse estadoResidencia = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(21).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(21).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = migrationFeign.findAllCitesByStateIdAndCountryId(bearerToken, paisResidencia.getId(), estadoResidencia.getId());
                    CountryResponse ciudadResidencia = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(22).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(22).getStringCellValue().concat(" not found"))));
                    direccionValues.put("Lugar de Residencia", Arrays.asList(paisResidencia, estadoResidencia, ciudadResidencia));

                    ProfileSecValueRequest contacto = new ProfileSecValueRequest();
                    contacto.setKeyword("PSCI06");
                    Map<String, Object> contactoValues = contacto.getFieldsValues();
                    contactoValues.put("Email Personal", row.getCell(17).getStringCellValue());
                    contactoValues.put("Número telefónico", (row.getCell(18) == null) ? "" : row.getCell(18).getStringCellValue());
                    contactoValues.put("Contacto 1", (row.getCell(23) == null) ? "" : row.getCell(23).getStringCellValue());

                    ProfileSecValueRequest dependientes = null;
                    if(row.getCell(26) != null) {
                        dependientes = new ProfileSecValueRequest();
                        dependientes.setKeyword("PSDP09");
                        Map<String, Object> dependientesValues = dependientes.getFieldsValues();
                        dependientesValues.put("Cantidad de dependientes económicos", (int) row.getCell(26).getNumericCellValue());
                    }

                    profileSecValueRequestList.add(informacionPersonal);
                    profileSecValueRequestList.add(informacionBiografica);
                    profileSecValueRequestList.add(datosPersonales);
                    profileSecValueRequestList.add(direccion);
                    profileSecValueRequestList.add(contacto);
                    if(dependientes != null) {
                        profileSecValueRequestList.add(dependientes);
                    }

                    profileRequest.setSectionValues(profileSecValueRequestList);

                    Long workPositionId = migrationFeign.findWorkPositionByCode(bearerToken, row.getCell(25).getStringCellValue()).getData().getWorkPosition().getId();

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
    public void migrateProfilesWorkPeriods(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("empleados");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    WorkPeriodAssignRequest workPeriodAssignRequest = new WorkPeriodAssignRequest();
                    workPeriodAssignRequest.setProfileIds(Collections.singleton(profileId));

                    Long workPeriodId = migrationFeign.findWorkPeriodByNameAndWorkPeriodType(bearerToken, row.getCell(16).getStringCellValue(), "Fixed Scheduled/Regular shift").getData().getId();

                    migrationFeign.createWorkPeriodAssignments(bearerToken, workPeriodAssignRequest, workPeriodId);

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

                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    ProfileSecValueRequest references = new ProfileSecValueRequest();
                    references.setKeyword("PSRF16");
                    Map<String, Object> referencesValues = references.getFieldsValues();
                    referencesValues.put("Nombre", (row.getCell(1) == null) ? "" : row.getCell(1).getStringCellValue());
                    referencesValues.put("Teléfono", (row.getCell(2) == null) ? "" : row.getCell(2).getStringCellValue());
                    referencesValues.put("Relación", (row.getCell(3) == null) ? "" : row.getCell(3).getStringCellValue());

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

                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    ProfileSecValueRequest informacionPago = new ProfileSecValueRequest();
                    informacionPago.setKeyword("PSPM14");
                    Map<String, Object> informacionPagoValues = informacionPago.getFieldsValues();
                    informacionPagoValues.put("Banco", row.getCell(1).getStringCellValue());
                    informacionPagoValues.put("Cuenta bancaria", row.getCell(2).getStringCellValue());
                    informacionPagoValues.put("Clabe interbancaria", row.getCell(3).getStringCellValue());
                    informacionPagoValues.put("Titular de la cuenta", row.getCell(4).getStringCellValue());

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
    public void migratePersonalLeaving(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("informacion de baja");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long profileId = migrationFeign.findProfileByClaveMpro(bearerToken, row.getCell(0).getStringCellValue()).getData().getId();

                    ProfileSecValueRequest personalLeaving = new ProfileSecValueRequest();
                    personalLeaving.setKeyword("PSPL18");
                    Map<String, Object> personalLeavingValues = personalLeaving.getFieldsValues();
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    String fechaBaja = "";
                    String fechaBajaSistema = "";
                    String fechaBajaHistorial = "";
                    if (row.getCell(1) != null) {
                        fechaBaja = row.getCell(1).getStringCellValue();
                        LocalDate.parse(fechaBaja, formatter);
                    }
                    if (row.getCell(2) != null) {
                        fechaBajaSistema = row.getCell(2).getStringCellValue();
                        LocalDate.parse(fechaBajaSistema, formatter);
                    }
                    if (row.getCell(3) != null) {
                        fechaBajaHistorial = row.getCell(3).getStringCellValue();
                        LocalDate.parse(fechaBajaHistorial, formatter);
                    }
                    personalLeavingValues.put("Fecha de baja", fechaBaja);
                    personalLeavingValues.put("Fecha de baja del sistema", fechaBajaSistema);
                    personalLeavingValues.put("Fecha de baja historial", fechaBajaHistorial); //no existe en bd
                    personalLeavingValues.put("Comentario de baja historial", (row.getCell(4) == null) ? "" : row.getCell(4).getStringCellValue()); //no existe en bd

                    migrationFeign.createProfileSectionValueByProfile(bearerToken, profileId, personalLeaving);
                } catch (ErrorResponseException e) {
                    log.error("Error processing row " + (i + 1) + " in sheet informacion de baja: " + e.getError().getErrors().getFields());

                    if(e.getError().getErrors().getId() != null) {
                        log.error("With model_fields id: " + e.getError().getErrors().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet informacion de baja: " + e.getMessage());
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

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheet("categorias de compensacion");

            logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new ConcurrentHashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if (columnName == null) {
                    cellCode = i;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("code")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("denomination")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("status")) {
                    cellStatus = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells());
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("Code column or denomination column do not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("Code column or denomination column do not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);
                    if(row.getCell(cellCode) == null) {
                        throw new NullCellException("Code cell can not be null");
                    }

                    if(row.getCell(cellDenomination) == null) {
                        throw new NullCellException("Denomination cell can not be null");
                    }

                    String code = (row.getCell(cellCode).getCellType() == CellType.STRING) ? (row.getCell(cellCode).getStringCellValue()).trim() : ("" + (int) row.getCell(cellCode).getNumericCellValue());
                    String denomination = (row.getCell(cellDenomination).getStringCellValue()).trim();
                    Map<String, Object> fieldsValues = new ConcurrentHashMap<>();

                    log.info("Compensacion a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por denominacion
                    DefaultResponse<List<CompCategoriesResponse>> compCategoriesRes = migrationFeign.simplifiedSearchCompensationCategories(bearerToken, denomination);
                    boolean existsCompCategoriesByDeno = compCategoriesRes.getData().stream()
                            .anyMatch(comp -> comp.getDenomination().equalsIgnoreCase(denomination));

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if(existsCompCategoriesByDeno) {
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    compCategoriesRes = migrationFeign.simplifiedSearchCompensationCategories(bearerToken, code);
                    boolean existsCompCategoriesbyCode = compCategoriesRes.getData().stream()
                            .anyMatch(comp -> comp.getCode().equalsIgnoreCase(code));

                    if (existsCompCategoriesbyCode) {
                        // Agregar celda con el mensaje de error en la fila que falló
                        Cell errorCell = row.createCell(row.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a compensation-category with the code");
                        continue;
                    }

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
                                        fieldsValues.put(nameColumn, cell.getNumericCellValue());
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

                    long idEstatus = 1L;
                    if(cellStatus != null) {
                        idEstatus = row.getCell(cellStatus).getStringCellValue().equalsIgnoreCase("active") ? 1L : 2L;
                    }

                    // Preparamos el objeto que irá en el body
                    CompCategoriesRequest compCategories = new CompCategoriesRequest();
                    compCategories.setCode(code);
                    compCategories.setDenomination(denomination);
                    compCategories.setFieldsValues(fieldsValues);
                    compCategories.setStatusId(idEstatus);

                    // Realizamos la petición
                    CompCategoriesResponse cPRes = migrationFeign.createCompensationCategories(bearerToken, compCategories).getData();
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    this.logRowErrorResponse(i, e);
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    this.logRowError(i, e);
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

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheet("tabuladores");

            this.logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new ConcurrentHashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;
            Integer cellMinSalary = null;
            Integer cellMaxSalary = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if (columnName == null) {
                    cellCode = i;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("code")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("denomination")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("status")) {
                    cellStatus = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("max_authorized_salary")) {
                    cellMinSalary = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("min_authorized_salary")) {
                    cellMaxSalary = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null || cellMinSalary == null || cellMaxSalary == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("code / denomination / min_salary / max_salary column do not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("code / denomination / min_salary / max_salary column do not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);

                    if(row.getCell(cellCode) == null) {
                        throw new NullCellException("code cell can not be null");
                    }

                    if(row.getCell(cellDenomination) == null) {
                        throw new NullCellException("denomination cell can not be null");
                    }

                    if(row.getCell(cellMinSalary) == null) {
                        throw new NullCellException("min_authorized_salary cell can not be null");
                    }

                    if(row.getCell(cellMaxSalary) == null) {
                        throw new NullCellException("max_authorized_salary cell can not be null");
                    }

                    // Sacamos el código y el nombre de la compensación
                    String code = (row.getCell(cellCode).getCellType() == CellType.STRING) ? (row.getCell(cellCode).getStringCellValue()).trim() : ("" + (int) row.getCell(cellCode).getNumericCellValue());
                    String denomination = (row.getCell(cellDenomination).getStringCellValue()).trim();
                    Map<String, Object> fieldsValues = new ConcurrentHashMap<>();
                    Long minSalary = (long) row.getCell(cellMinSalary).getNumericCellValue();
                    Long maxSalary = (long) row.getCell(cellMaxSalary).getNumericCellValue();

                    log.info("Tabulador a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por nombre
                    DefaultResponse<List<TabsResponse>> tabsRes = migrationFeign.simplifiedSearchTabs(bearerToken, denomination);
                    boolean existsTabByDeno = tabsRes.getData().stream()
                            .anyMatch(comp -> comp.getDenomination().equalsIgnoreCase(denomination));

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsTabByDeno) {
                        log.info("Continuamos debido a que el tabulador ya existe!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    tabsRes = migrationFeign.simplifiedSearchTabs(bearerToken, code);
                    boolean existsTabByCode = tabsRes.getData().stream()
                            .anyMatch(comp -> comp.getCode().equalsIgnoreCase(code));

                    if (existsTabByCode) {

                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a compensation-tab with the code");
                        continue;
                    }

                    if (minSalary < 0 || maxSalary < 0) {
                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: max_authorized_salary or min_authorized_salary can not be less than zero");
                        continue;
                    }

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
                                        fieldsValues.put(nameColumn, cell.getNumericCellValue());
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

                    long idEstatus = 1L;
                    if(cellStatus != null) {
                        idEstatus = row.getCell(cellStatus).getStringCellValue().equalsIgnoreCase("active") ? 1L : 2L;
                    }

                    // Preparamos el objeto que irá en el body
                    TabsRequest tabsRequest = new TabsRequest();
                    tabsRequest.setCode(code);
                    tabsRequest.setDenomination(denomination);
                    tabsRequest.setMinAuthorizedSalary(minSalary);
                    tabsRequest.setMaxAuthorizedSalary(maxSalary);
                    tabsRequest.setStatusId(idEstatus);
                    tabsRequest.setFieldsValues(fieldsValues);

                    // Realizamos la petición
                    TabsResponse tabRes = migrationFeign.createTab(bearerToken, tabsRequest).getData();
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    this.logRowErrorResponse(i, e);
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    this.logRowError(i, e);
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
            // Nos posicionamos en la hoja
            Sheet sheet = workbook.getSheet("puestos");

            this.logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Map<String, Integer> fieldsExcel = new ConcurrentHashMap<>();
            Integer cellCode = null;
            Integer cellDenomination = null;
            Integer cellStatus = null;

            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                
                Cell columnName = rowNames.getCell(i);
                if(columnName == null) {
                    continue;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("code")) {
                    cellCode = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("denomination")) {
                    cellDenomination = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("status")) {
                    cellStatus = i;
                } else {
                    fieldsExcel.put(columnName.getStringCellValue(), i);
                }
            }

            if(cellCode == null || cellDenomination == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("code / denomination column do not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("code / denomination column do not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {
                    Row row = sheet.getRow(i);

                    if(row.getCell(cellCode) == null) {
                        throw new NullCellException("code cell can not be null");
                    }

                    if(row.getCell(cellDenomination) == null) {
                        throw new NullCellException("denomination cell can not be null");
                    }

                    // Sacamos el código y el nombre de la compensación
                    String code = (row.getCell(cellCode).getCellType() == CellType.STRING) ? (row.getCell(cellCode).getStringCellValue()).trim() : ("" + (int) row.getCell(cellCode).getNumericCellValue());
                    String denomination = (row.getCell(cellDenomination).getStringCellValue()).trim();
                    Map<String, Object> fieldsValues = new ConcurrentHashMap<>();

                    log.info("Puesto a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por nombre
                    DefaultResponse<List<WorkPositionCategoryResponse>> worksPositionsCategoriesRes = migrationFeign.simplifiedSearchWorkPositionCategory(bearerToken, denomination);
                    boolean existsWorkByDeno = worksPositionsCategoriesRes.getData().stream()
                            .anyMatch(comp -> {
                                if(comp.getDenomination().equalsIgnoreCase(denomination)) {
                                    return true;
                                }
                                return false;
                            });

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsWorkByDeno) {
                        log.info("Continuamos debido a que ese puesto ya existe!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    worksPositionsCategoriesRes = migrationFeign.simplifiedSearchWorkPositionCategory(bearerToken, code);
                    boolean existsWorkByCode = worksPositionsCategoriesRes.getData().stream()
                            .anyMatch(comp -> comp.getCode().equalsIgnoreCase(code));

                    if (existsWorkByCode) {
                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a work-positions-category with the code");
                        continue;
                    }

                    fieldsExcel.forEach((nameColumn, position) -> {
                        Cell cell = row.getCell(position);
                        log.info(nameColumn);
                        if (cell == null) {
                            fieldsValues.put(nameColumn, "");
                        } else {
                            switch (cell.getCellType()) {
                                case STRING:
                                    fieldsValues.put(nameColumn, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        fieldsValues.put(nameColumn, cell.getDateCellValue());
                                    } else {
                                        fieldsValues.put(nameColumn, (int) cell.getNumericCellValue());
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

                    long idEstatus = 1L;
                    if(cellStatus != null) {
                        idEstatus = row.getCell(cellStatus).getStringCellValue().equalsIgnoreCase("active") ? 1L : 2L;
                    }

                    // Preparamos el objeto que irá en el body
                    WorkPositionCategoryRequest workPositionCategoryRequest = new WorkPositionCategoryRequest();
                    workPositionCategoryRequest.setCode(code);
                    workPositionCategoryRequest.setDenomination(denomination);
                    workPositionCategoryRequest.setFieldsValues(fieldsValues);
                    workPositionCategoryRequest.setStatusId(idEstatus);

                    // Realizamos la petición
                    DefaultResponse<WorkPositionCategoryResponse> wpc = migrationFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    this.logRowErrorResponse(i, e);
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    this.logRowError(i, e);
                    this.agregarCeldaError(sheet.getRow(i), e.getMessage());
                }
            }
            modifiedFile = this.createModifiedWorkbook(workbook, file);
        } catch (Exception e) {
            this.logProcessingExcelFile(e);
        }
        return modifiedFile;
    }

    public File loadWorkPeriods(MultipartFile file) {

        String bearerToken = this.getBearerToken();

        List<WorkPeriodTypeResponse> workPeriodTypesList = migrationFeign.findAllWorkPeriodTypes(bearerToken).getData();
        List<WorkPeriodMaxDurationsResponse> workPeriodMaxDurationsList = migrationFeign.findAllWorkPeriodsMaxDurations(bearerToken).getData();
        List<WorkPeriodMaxDailyDurationsResponse> workPeriodMaxDailyDurationsResponseList = migrationFeign.findAllWorkPeriodsMaxDailyDurations(bearerToken).getData();
        List<DurationsResponse> durations = migrationFeign.findAllDurations(bearerToken).getData();
        List<WorkTurnTypesResponse> workturntypes = migrationFeign.findAllWorkTurnTypes(bearerToken).getData();

        // Archivo modificado para devolver
        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheet("jornadas");

            this.logSheetNameNumberOfRows(sheet);

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = this.greenCellStyle(workbook);

            Row rowNames = sheet.getRow(0);
            Integer cellName = null;
            Integer cellPeriodType = null;
            Integer cellKeywordMaxDuration = null;
            Integer cellMaxDailyDuration = null;


            for (int i = 0; i < rowNames.getPhysicalNumberOfCells(); i++) {
                Cell columnName = rowNames.getCell(i);

                if(columnName == null) {
                    continue;
                } else if (columnName.getStringCellValue().equalsIgnoreCase("name")) {
                    cellName = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("period_type")) {
                    cellPeriodType = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("keyword_max_duration")) {
                    cellKeywordMaxDuration = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("max_daily_duration")) {
                    cellMaxDailyDuration = i;
                }
            }

            if(cellName == null || cellPeriodType == null || cellKeywordMaxDuration == null || cellMaxDailyDuration == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("name / period_type / keyword_max_duration / max_daily_duration column do not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("name / period_type / keyword_max_duration / max_daily_duration column do not exist");
            }

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                try {

                    List<WorkPeriodDetailRequest> workPeriodDetailList = new ArrayList<>();

                    Row row = sheet.getRow(i);

                    if(row.getCell(cellName) == null) {
                        throw new NullCellException("name cell can not be null");
                    }

                    if(row.getCell(cellPeriodType) == null) {
                        throw new NullCellException("period_type cell can not be null");
                    }

                    if(row.getCell(cellKeywordMaxDuration) == null) {
                        throw new NullCellException("keyword_max_duration cell can not be null");
                    }

                    String name = (row.getCell(cellName).getStringCellValue()).trim();
                    String periodType = (row.getCell(cellPeriodType).getStringCellValue()).trim();
                    String keywordMaxDuration = (row.getCell(cellKeywordMaxDuration).getStringCellValue()).trim();
                    Integer maxDailyDuration = (row.getCell(cellMaxDailyDuration) == null) ? null : (int)row.getCell(cellMaxDailyDuration).getNumericCellValue();

                    log.info("Periodo de trabajo a consultar con nombre: " + name);

                    // Consultamos si existe el periodo de trabajo por nombre
                    // DefaultResponse<WorkPeriodResponse> workPeriodResponse = workPeriodsFeign.findOneByName(bearerToken, name);
                    // boolean existsWorkPeriod = workPeriodResponse.getData() != null &&
                    //         workPeriodResponse.getData().getName().equalsIgnoreCase(name);

                    // // Si existe, seguimos a la siguiente para no volverla a insertar
                    // if (existsWorkPeriod) {
                    //     log.info("Continuamos debido a que existe una jornada laboral con ese nombre!");
                    //     workPeriodsMap.put(workPeriodResponse.getData().getName(), workPeriodResponse.getData().getId());
                    //     row.getCell(cellName).setCellStyle(cellStyle);
                    //     continue;
                    // }

                    Long idWorkPeriodType;
                    Long idWorkPeriodMaxDailyDuration = null;
                    if(periodType.equalsIgnoreCase("Horario Fijo")) {

                        idWorkPeriodType = workPeriodTypesList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("Fixed Scheduled/Regular shift"))
                                .map(WorkPeriodTypeResponse::getId).findFirst().get();

                        log.info("El valor de max daily duration es: " + maxDailyDuration);

                        idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpmd -> wpmd.getDuration() == maxDailyDuration)
                                    .map(WorkPeriodMaxDailyDurationsResponse::getId).findFirst().orElseThrow(() -> new Exception("There is no max_daily_duration"));

                    } else if(periodType.equalsIgnoreCase("Frecuencia Variable")) {

                        idWorkPeriodType = workPeriodTypesList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("Variable frecuency shift"))
                                .map(WorkPeriodTypeResponse::getId).findFirst().orElseThrow();

                    } else {
                        throw new NullCellException("There is no period with that name");
                    }

                    Long idWorkPeriodMaxDuration = workPeriodMaxDurationsList
                            .stream()
                            .filter(wpmd -> wpmd.getKeyword().equalsIgnoreCase(keywordMaxDuration))
                            .map(WorkPeriodMaxDurationsResponse::getId).findFirst().orElseThrow(() -> new Exception("There is no max_duration with that keyword"));

                    // Nos posicionamos en la segunda hoja donde estan los tunos de trabajo
                    Sheet workTurnsSheet = workbook.getSheet("turnos");

                    this.logSheetNameNumberOfRows(workTurnsSheet);

                    // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
                    for (int j = 1; j < workTurnsSheet.getPhysicalNumberOfRows(); j++) {

                        try {
                            Row workTurnRow = workTurnsSheet.getRow(j);

                            if(workTurnRow.getCell(0) == null) {
                                throw new NullCellException("work_period name is null");
                            }

                            String nameWorkPeriod = workTurnRow.getCell(0).getStringCellValue();
                            Date dateFrom = null;
                            Date dateTo = null;
                            String from = "";
                            String to = "";

                            //Verificamos si el nombre coincide sino continuamos
                            if(!nameWorkPeriod.equalsIgnoreCase(name)) {
                                continue;
                            }

                            if(periodType.equalsIgnoreCase("Horario Fijo")) {

                                if(workTurnRow.getCell(1) == null) {
                                    throw new NullCellException("date_from is null");
                                }
                                
                                if(workTurnRow.getCell(2) == null) {
                                    throw new NullCellException("date_to is null");
                                }

                                dateFrom = workTurnRow.getCell(1).getDateCellValue();
                                dateTo = workTurnRow.getCell(2).getDateCellValue();
                            }

                            if(workTurnRow.getCell(3) == null) {
                                throw new NullCellException("day_of_week name is null");
                            }

                            if(workTurnRow.getCell(4) == null) {
                                throw new NullCellException("work_turn_type is null");
                            }

                            Integer dayOfWeek = (int) workTurnRow.getCell(3).getNumericCellValue();
                            String workTurnType = workTurnRow.getCell(4).getStringCellValue();
                            Integer duration = (workTurnRow.getCell(5) == null) ? 0 : (int) workTurnRow.getCell(5).getNumericCellValue();

                            Long idWorkTurnType = workturntypes
                                    .stream()
                                    .filter(wtt -> wtt.getName().equalsIgnoreCase(workTurnType))
                                    .map(wtt -> wtt.getId()).findFirst().orElseThrow();

                            Long idDuration = null;
                            if(duration != 0){
                                idDuration = durations
                                        .stream()
                                        .filter(d -> d.getAmount() == duration)
                                        .map(d -> d.getId()).findFirst().orElseThrow();
                            }

                            // Formatear la hora
                            SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm");

                            if(dateFrom != null && dateTo != null) {
                                from = timeFormat.format(dateFrom);
                                to = timeFormat.format(dateTo);
                            }

                            workPeriodDetailList.add(new WorkPeriodDetailRequest(
                                    from, to, dayOfWeek,
                                    idWorkTurnType, idDuration));

                            workTurnRow.getCell(0).setCellStyle(cellStyle);
                        } catch (Exception e) {
                            this.agregarCeldaError(workTurnsSheet.getRow(j), e.getMessage());
                        }
                    }

                    // Preparamos el objeto que irá en el body
                    WorkPeriodRequest workPeriodRequest = new WorkPeriodRequest();
                    workPeriodRequest.setName(name);
                    workPeriodRequest.setWorkPeriodTypeId(idWorkPeriodType);
                    workPeriodRequest.setWorkTurns(workPeriodDetailList);
                    workPeriodRequest.setWorkTurns(workPeriodDetailList);
                    workPeriodRequest.setWorkPeriodMaxDurationId(idWorkPeriodMaxDuration);
                    workPeriodRequest.setWorkPeriodMaxDailyDurationId(idWorkPeriodMaxDailyDuration);

                    log.info("LLenamos todos los workPeriod con el id nesario: \nName: " + workPeriodRequest.getName()
                            + "\nTipo de periodo de trabajo: " + idWorkPeriodType
                            + "\nMaximo de duracion id: " + idWorkPeriodMaxDuration
                            + "\nMaximo de duracion por dia id: " + idWorkPeriodMaxDailyDuration);

                    // Realizamos la petición
                    DefaultResponse<WorkPeriodResponse> wpr = migrationFeign.createWorkPeriods(bearerToken, workPeriodRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch(ErrorResponseException e) {
                    this.logRowErrorResponse(i, e);
                    ErrorResponse error = e.getError();
                    this.agregarExcetionFeign(sheet.getRow(i), error.getErrors().getFields());
                } catch (Exception e) {
                    this.logRowError(i, e);
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
                    this.logRowErrorResponse(i, e);
                } catch (Exception e) {
                    this.logRowError(i, e);
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

    private void logRowError(int i, Exception e) {
        log.error("Error processing row " + (i + 1) + " in Excel: " + e.getMessage());
    }

    private void logRowErrorResponse(int i, ErrorResponseException e) {
        log.error("Error processing row " + (i + 1) + " in Excel: " + e.getError().getErrors().getFields().toString());
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
