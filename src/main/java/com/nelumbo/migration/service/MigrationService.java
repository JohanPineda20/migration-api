package com.nelumbo.migration.service;

import com.nelumbo.migration.exceptions.ErrorResponseException;
import com.nelumbo.migration.exceptions.NullCellException;
import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;

import feign.FeignException;
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

@Slf4j
@Service
@RequiredArgsConstructor
public class MigrationService {

    private final CountryFeign countryFeign;
    private final CostCenterFeign costCenterFeign;
    private final LoginFeign loginFeign;
    private final StoreFeign storeFeign;
    private final OrgEntityFeign orgEntityFeign;
    private final WorkPositionFeign workPositionFeign;
    private final ProfileFeign profileFeign;
    private final CompCategoriesFeign compCategoriesFeign;
    private final TabsFeign tabsFeign;
    private final WorksPositionCategoriesFeign worksPositionCategoriesFeign;
    private final WorkPeriodsFeign workPeriodsFeign;
    private final WorkPeriodsTypesFeign workPeriodsTypesFeign;
    private final WorkPeriodsMaxDurationsFeign workPeriodsMaxDurationsFeign;
    private final WorkPeriodsMaxDailyDurationsFeign workPeriodsMaxDailyDurationsFeign;
    private final DurationsFeign durationsFeign;
    private final WorkTurnTypesFeign workTurnTypesFeign;

    Map<String, Long> compensationCategoriesResponseMap = new ConcurrentHashMap<>();
    Map<String, Long> compensationTabResponseMap = new ConcurrentHashMap<>();
    Map<String, Long> costCenterResponseMap = new ConcurrentHashMap<>();
    Map<String, Long> storeResponseMap = new ConcurrentHashMap<>();
    Map<String,Map<String, Long>> storeDetailResponseMap = new ConcurrentHashMap<>();
    Map<String,Long> workPositionResponseMap = new ConcurrentHashMap<>();
    Map<String,Long> workPeriodsMap = new ConcurrentHashMap<>();
    Map<String,Long> workPositionCategoriesMap = new ConcurrentHashMap<>();
    
    @Value("${email}")
    private String email;

    @Value("${password}")
    private String password;

    //constantes
    private static final String BEARER = "Bearer ";
    private static final String MODIFIED = "modified_";
    private static final String SHEET = "Estamos con la hoja: ";
    private static final String COUNTROWS = "La cantidad de filas es: ";

    public void migrateCostCenters(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("centro de costos");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);
                    CostCenterRequest costCenterRequest = new CostCenterRequest();
                    Cell cellCode = row.getCell(0);
                    costCenterRequest.setCode(cellCode.getCellType() == CellType.STRING ? cellCode.getStringCellValue() : "" + (int) cellCode.getNumericCellValue());
                    costCenterRequest.setDenomination(row.getCell(1).getStringCellValue());

                    DefaultResponse<List<CountryResponse>> countryResponse = countryFeign.findAll();
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(2).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(2).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = countryFeign.findAllStatesByCountryId(countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(3).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(3).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = countryFeign.findAllCitesByStateIdAndCountryId(countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(4).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(4).getStringCellValue().concat(" not found"))));

                    costCenterRequest.setCountryId(countryId);
                    costCenterRequest.setStateId(stateId);
                    costCenterRequest.setCityId(cityId);
                    costCenterRequest.setStatusId(1L);

                    DefaultResponse<CostCenterResponse> costCenterResponse = costCenterFeign.createCostCenter(bearerToken, costCenterRequest);
                    costCenterResponseMap.put(costCenterResponse.getData().getDenomination(), costCenterResponse.getData().getId());
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
                    storeRequest.setCode(code.getCellType() == CellType.STRING ? code.getStringCellValue() : "" + (int) code.getNumericCellValue());
                    storeRequest.setDenomination(row.getCell(1).getStringCellValue());

                    DefaultResponse<List<CountryResponse>> countryResponse = countryFeign.findAll();
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(2).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(2).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = countryFeign.findAllStatesByCountryId(countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(3).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(3).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = countryFeign.findAllCitesByStateIdAndCountryId(countryId, stateId);
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
                    Long costCenterId = row.getCell(10) != null ? costCenterResponseMap.get(row.getCell(10).getStringCellValue()) : null;
                    storeRequest.setCostCenterId(costCenterId);
                    DefaultResponse<StoreResponse> storeResponse = storeFeign.createStore(bearerToken, storeRequest);
                    storeResponseMap.put(storeResponse.getData().getDenomination(), storeResponse.getData().getId());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursales: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateStoresOrgEntities(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("sucursal_org_entities");
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try {
                    Row row = sheet.getRow(i);

                    Long storeId = storeResponseMap.get(row.getCell(0).getStringCellValue());

                    StoreDetailRequest storeDetailRequest = new StoreDetailRequest();
                    List<Long> orgEntityDetailIds = storeDetailRequest.getOrgEntityDetailIds();

                    Long regionId = null;
                    Long divisionId = null;
                    Long zonaId = null;

                    Cell cellRegion = row.getCell(1);
                    Cell cellDivision = row.getCell(2);
                    Cell cellZona = row.getCell(3);

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
                    storeFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);

                    Cell cellDepartamento = row.getCell(4);
                    if (cellDepartamento == null) throw new RuntimeException("store need one or more departments");
                    String [] departments = cellDepartamento.getStringCellValue().split(",");

                    storeDetailResponseMap.put(row.getCell(0).getStringCellValue(), new HashMap<>());
                    for (String department : departments) {
                        DefaultResponse<Page<OrgEntityResponse>> entityResponse = orgEntityFeign.findAllInstancesParentOrganizationEntityDetail(
                                bearerToken, 5L, 0L);
                        Long departmentId = entityResponse.getData().getContent().stream()
                                .filter(entity -> entity.getName().equalsIgnoreCase(department))
                                .findFirst()
                                .map(OrgEntityResponse::getId)
                                .orElseThrow(() -> new RuntimeException("department ".concat(department).concat(" not found")));
                        orgEntityDetailIds = new ArrayList<>();
                        orgEntityDetailIds.add(departmentId);
                        storeDetailRequest.setOrgEntityDetailIds(orgEntityDetailIds);
                        DefaultResponse<StoreDetailResponse> storeDetailResponse = storeFeign.createStoreDetails(bearerToken, storeDetailRequest, storeId);
                        storeDetailResponseMap.get(row.getCell(0).getStringCellValue()).put(department, storeDetailResponse.getData().getId());
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_org_entities: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    private Long getEntityId(String bearerToken, Cell cell, Long entityType, Long parentId, String entityName) {
        DefaultResponse<Page<OrgEntityResponse>> entityResponse = orgEntityFeign.findAllInstancesParentOrganizationEntityDetail(
                bearerToken, entityType, parentId
        );

        return entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(cell.getStringCellValue()))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(cell.getStringCellValue()).concat(" not found")));
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
                    workPositionRequest.setCode(code.getCellType() == CellType.STRING ? code.getStringCellValue() : "" + (int) code.getNumericCellValue());
                    workPositionRequest.setDenomination(row.getCell(1).getStringCellValue());
                    workPositionRequest.setAuthorizedStaff((long)row.getCell(2).getNumericCellValue());
                    workPositionRequest.setStatusId(1L);
                    workPositionRequest.setWorkPosCatId(workPositionCategoriesMap.get(row.getCell(3).getStringCellValue()));
                    workPositionRequest.setStoreId(storeResponseMap.get(row.getCell(4).getStringCellValue()));
                    workPositionRequest.setStoreOrganizativeId(storeDetailResponseMap.get(row.getCell(4).getStringCellValue()).get(row.getCell(5).getStringCellValue()));
                    Long costCenterId = row.getCell(6) != null ? costCenterResponseMap.get(row.getCell(6).getStringCellValue()) : null;
                    workPositionRequest.setCostCenterId(costCenterId);
                    DefaultResponse<WorkPositionDetailResponse> workPositionDetailResponse = workPositionFeign.createWorkPosition(bearerToken, workPositionRequest);
                    workPositionResponseMap.put(workPositionDetailResponse.getData().getWorkPosition().getDenomination(), workPositionDetailResponse.getData().getWorkPosition().getId());
                    
                    WorkPositionUpdateRequest wPUReq = WorkPositionUpdateRequest.builder()
                            .compCategoryId(compensationCategoriesResponseMap.get(row.getCell(7).getStringCellValue()))
                            .compTabId(compensationTabResponseMap.get(row.getCell(8).getStringCellValue()))
                            .minSalary((long)row.getCell(9).getNumericCellValue())
                            .build();

                    workPositionFeign.updateWorkPosition(bearerToken, wPUReq, workPositionDetailResponse.getData().getWorkPosition().getId());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargos: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateProfiles(MultipartFile file) {
        String bearerToken = this.getBearerToken();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheet("perfiles");
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
                    informacionPersonalValues.put("Segundo Nombre", row.getCell(2).getStringCellValue());
                    informacionPersonalValues.put("Primer Apellido", row.getCell(3).getStringCellValue());
                    informacionPersonalValues.put("Segundo Apellido", row.getCell(4).getStringCellValue());
                    informacionPersonalValues.put("Sexo", row.getCell(5).getStringCellValue());
                    informacionPersonalValues.put("Grupo sanguíneo", row.getCell(12).getStringCellValue());
                    informacionPersonalValues.put("Estado civil", row.getCell(14).getStringCellValue());
                    DateTimeFormatter formatters = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    LocalDate.parse(row.getCell(15).getStringCellValue(), formatters);
                    informacionPersonalValues.put("Fecha de contratación", row.getCell(15).getStringCellValue());
                    informacionPersonalValues.put("Clave antigua", clave.getCellType() == CellType.STRING ? clave.getStringCellValue() : "" + (int) clave.getNumericCellValue());

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

                    ProfileSecValueRequest direccion = new ProfileSecValueRequest();
                    direccion.setKeyword("PSAS05");
                    Map<String, Object> direccionValues = direccion.getFieldsValues();
                    direccionValues.put("Dirección", row.getCell(19).getStringCellValue());
                    DefaultResponse<List<CountryResponse>> countryResponse = countryFeign.findAll();
                    CountryResponse paisResidencia = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(20).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("country ".concat(row.getCell(2).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> stateResponse = countryFeign.findAllStatesByCountryId(paisResidencia.getId());
                    CountryResponse estadoResidencia = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(21).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("state ".concat(row.getCell(3).getStringCellValue().concat(" not found"))));
                    DefaultResponse<List<CountryResponse>> cityResponse = countryFeign.findAllCitesByStateIdAndCountryId(paisResidencia.getId(), estadoResidencia.getId());
                    CountryResponse ciudadResidencia = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(22).getStringCellValue()))
                            .findFirst().orElseThrow(() -> new RuntimeException("city ".concat(row.getCell(4).getStringCellValue().concat(" not found"))));
                    direccionValues.put("Lugar de Residencia", Arrays.asList(paisResidencia, estadoResidencia, ciudadResidencia));

                    ProfileSecValueRequest contacto = new ProfileSecValueRequest();
                    contacto.setKeyword("PSCI06");
                    Map<String, Object> contactoValues = contacto.getFieldsValues();
                    contactoValues.put("Email Personal", row.getCell(17).getStringCellValue());
                    contactoValues.put("Número telefónico", (int) row.getCell(18).getNumericCellValue());

                    profileSecValueRequestList.add(informacionPersonal);
                    profileSecValueRequestList.add(informacionBiografica);
                    profileSecValueRequestList.add(datosPersonales);
                    profileSecValueRequestList.add(direccion);
                    profileSecValueRequestList.add(contacto);
                    profileRequest.setSectionValues(profileSecValueRequestList);
                    Long workPositionId = workPositionResponseMap.get(row.getCell(25).getStringCellValue());
                    if (workPositionId == null) throw new RuntimeException("work position ".concat(row.getCell(25).getStringCellValue().concat(" not found")));
                    profileRequest.setWorkPositionId(workPositionId);
                    DefaultResponse<ProfileResponse> profileResponse = profileFeign.createProfile(bearerToken, profileRequest);

                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet perfiles: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
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
                    String [] jornadas = row.getCell(1).getStringCellValue().split(",");
                    Long storeId = storeResponseMap.get(row.getCell(0).getStringCellValue());
                    if (storeId == null) throw new RuntimeException("Store ".concat(row.getCell(0).getStringCellValue()).concat(" not found"));
                    for (String jornada:jornadas) {
                        storeWorkPeriodRequest.setWorkPeriodId(workPeriodsMap.get(jornada));
                        storeFeign.createStoreWorkPeriods(bearerToken, storeWorkPeriodRequest, storeId);
                    }
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet sucursal_jornadas: " + e.getMessage());
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
            Sheet sheet = workbook.getSheet("compensation_categories");

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
                    DefaultResponse<List<CompCategoriesResponse>> compCategoriesRes = compCategoriesFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsCompCategoriesByDeno = compCategoriesRes.getData().stream()
                            .anyMatch(comp -> comp.getDenomination().equalsIgnoreCase(denomination));

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if(existsCompCategoriesByDeno) {
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    compCategoriesRes = compCategoriesFeign.simplifiedSearch(bearerToken, code);
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
                    CompCategoriesResponse cPRes = compCategoriesFeign.createCompensationCategories(bearerToken, compCategories).getData();
                    this.compensationCategoriesResponseMap.put(cPRes.getCode(), cPRes.getId());
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
            Sheet sheet = workbook.getSheet("compensation_tab");

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
                    DefaultResponse<List<TabsResponse>> tabsRes = tabsFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsTabByDeno = tabsRes.getData().stream()
                            .anyMatch(comp -> comp.getDenomination().equalsIgnoreCase(denomination));

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsTabByDeno) {
                        log.info("Continuamos debido a que el tabulador ya existe!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    tabsRes = tabsFeign.simplifiedSearch(bearerToken, code);
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
                    TabsResponse tabRes = tabsFeign.createTab(bearerToken, tabsRequest).getData();
                    compensationTabResponseMap.put(tabRes.getCode(), tabRes.getId());
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
            Sheet sheet = workbook.getSheet("work_position_categories");

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
                    DefaultResponse<List<WorkPositionCategoryResponse>> worksPositionsCategoriesRes = worksPositionCategoriesFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsWorkByDeno = worksPositionsCategoriesRes.getData().stream()
                            .anyMatch(comp -> {
                                if(comp.getDenomination().equalsIgnoreCase(denomination)) {
                                    workPositionCategoriesMap.put(comp.getDenomination(), comp.getId());
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
                    worksPositionsCategoriesRes = worksPositionCategoriesFeign.simplifiedSearch(bearerToken, code);
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
                    WorkPositionCategoryRequest workPositionCategoryRequest = new WorkPositionCategoryRequest();
                    workPositionCategoryRequest.setCode(code);
                    workPositionCategoryRequest.setDenomination(denomination);
                    workPositionCategoryRequest.setFieldsValues(fieldsValues);
                    workPositionCategoryRequest.setStatusId(idEstatus);

                    // Realizamos la petición
                    DefaultResponse<WorkPositionCategoryResponse> wpc = worksPositionCategoriesFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    workPositionCategoriesMap.put(wpc.getData().getDenomination(), wpc.getData().getId());
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

        List<WorkPeriodTypeResponse> workPeriodTypesList = workPeriodsTypesFeign.findAllWorkPeriodTypes(bearerToken).getData();
        List<WorkPeriodMaxDurationsResponse> workPeriodMaxDurationsList = workPeriodsMaxDurationsFeign.findAllWorkPeriodsMaxDurations(bearerToken).getData();
        List<WorkPeriodMaxDailyDurationsResponse> workPeriodMaxDailyDurationsResponseList = workPeriodsMaxDailyDurationsFeign.findAllWorkPeriodsMaxDailyDurations(bearerToken).getData();
        List<DurationsResponse> durations = durationsFeign.findAllDurations(bearerToken).getData();
        List<WorkTurnTypesResponse> workturntypes = workTurnTypesFeign.findAllWorkTurnTypes(bearerToken).getData();

        // Archivo modificado para devolver
        File modifiedFile = new File(MODIFIED + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheet("work_periods");

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
                } else if(columnName.getStringCellValue().equalsIgnoreCase("keyword_max_duracion")) {
                    cellKeywordMaxDuration = i;
                } else if(columnName.getStringCellValue().equalsIgnoreCase("max_daily_duration")) {
                    cellMaxDailyDuration = i;
                }
            }

            if(cellName == null || cellPeriodType == null || cellKeywordMaxDuration == null || cellMaxDailyDuration == null) {
                Cell cell = rowNames.createCell(rowNames.getPhysicalNumberOfCells() + 1);
                cell.setCellStyle(this.redCellStyle(workbook));
                cell.setCellValue("name / period_type / keyword_max_duracion / max_daily_duration column do not exist");
                modifiedFile = this.createModifiedWorkbook(workbook, file);
                throw new NullCellException("name / period_type / keyword_max_duracion / max_daily_duration column do not exist");
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
                        throw new NullCellException("keyword_max_duracion cell can not be null");
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
                    Sheet workTurnsSheet = workbook.getSheet("work_turns");

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
                    DefaultResponse<WorkPeriodResponse> wpr = workPeriodsFeign.createWorkPeriods(bearerToken, workPeriodRequest);
                    workPeriodsMap.put(wpr.getData().getName(), wpr.getData().getId());
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
