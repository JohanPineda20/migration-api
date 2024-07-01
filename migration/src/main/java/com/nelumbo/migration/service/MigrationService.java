package com.nelumbo.migration.service;

import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

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
    Map<String, Long> costCenterResponseMap = new ConcurrentHashMap<>();
    Map<String, Long> storeResponseMap = new ConcurrentHashMap<>();
    Map<String,Map<String, Long>> storeDetailResponseMap = new ConcurrentHashMap<>();
    Map<String,Long> workPositionResponseMap = new ConcurrentHashMap<>();
    @Value("${email}")
    private String email;
    @Value("${password}")
    private String password;

    public void migrateCostCenters(MultipartFile file) {
        DefaultResponse<LoginResponse> loginResponse = login();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(0);
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

                    DefaultResponse<CostCenterResponse> costCenterResponse = costCenterFeign.createCostCenter("Bearer ".concat(loginResponse.getData().getToken()), costCenterRequest);
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
        DefaultResponse<LoginResponse> loginResponse = login();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(1);
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
                    DefaultResponse<StoreResponse> storeResponse = storeFeign.createStore("Bearer ".concat(loginResponse.getData().getToken()), storeRequest);
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
        DefaultResponse<LoginResponse> loginResponse = login();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(2);
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
                            regionId = getEntityId(loginResponse, cellRegion, 2L, 1L, "region");
                            orgEntityDetailIds.add(regionId);
                        }

                        if (cellDivision != null) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region");
                            }
                            divisionId = getEntityId(loginResponse, cellDivision, 3L, regionId, "division");
                            orgEntityDetailIds.add(divisionId);
                        }

                        if (cellZona != null) {
                            if (regionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing region and division");
                            }
                            if (divisionId == null) {
                                throw new RuntimeException("Invalid geographic structure: missing division");
                            }
                            zonaId = getEntityId(loginResponse, cellZona, 4L, divisionId, "zona");
                            orgEntityDetailIds.add(zonaId);
                        }
                    }
                    storeFeign.createStoreDetails("Bearer ".concat(loginResponse.getData().getToken()), storeDetailRequest, storeId);

                    Cell cellDepartamento = row.getCell(4);
                    if (cellDepartamento == null) throw new RuntimeException("store need one or more departments");
                    String [] departments = cellDepartamento.getStringCellValue().split(",");

                    storeDetailResponseMap.put(row.getCell(0).getStringCellValue(), new HashMap<>());
                    for (String department : departments) {
                        DefaultResponse<Page<OrgEntityResponse>> entityResponse = orgEntityFeign.findAllInstancesParentOrganizationEntityDetail(
                                "Bearer ".concat(loginResponse.getData().getToken()), 5L, 0L);
                        Long departmentId = entityResponse.getData().getContent().stream()
                                .filter(entity -> entity.getName().equalsIgnoreCase(department))
                                .findFirst()
                                .map(OrgEntityResponse::getId)
                                .orElseThrow(() -> new RuntimeException("department ".concat(department).concat(" not found")));
                        orgEntityDetailIds = new ArrayList<>();
                        orgEntityDetailIds.add(departmentId);
                        storeDetailRequest.setOrgEntityDetailIds(orgEntityDetailIds);
                        DefaultResponse<StoreDetailResponse> storeDetailResponse = storeFeign.createStoreDetails("Bearer ".concat(loginResponse.getData().getToken()), storeDetailRequest, storeId);
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

    private Long getEntityId(DefaultResponse<LoginResponse> loginResponse, Cell cell, Long entityType, Long parentId, String entityName) {
        DefaultResponse<Page<OrgEntityResponse>> entityResponse = orgEntityFeign.findAllInstancesParentOrganizationEntityDetail(
                "Bearer ".concat(loginResponse.getData().getToken()), entityType, parentId
        );

        return entityResponse.getData().getContent().stream()
                .filter(entity -> entity.getName().equalsIgnoreCase(cell.getStringCellValue()))
                .findFirst()
                .map(OrgEntityResponse::getId)
                .orElseThrow(() -> new RuntimeException(entityName.concat(" ").concat(cell.getStringCellValue()).concat(" not found")));
    }

    public void migrateWorkPositions(MultipartFile file) {
        DefaultResponse<LoginResponse> loginResponse = login();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(3);
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
                    workPositionRequest.setWorkPosCatId(1L);//workPosCatResponseMap.get(row.getCell(3).getStringCellValue())
                    workPositionRequest.setStoreId(storeResponseMap.get(row.getCell(4).getStringCellValue()));
                    workPositionRequest.setStoreOrganizativeId(storeDetailResponseMap.get(row.getCell(4).getStringCellValue()).get(row.getCell(5).getStringCellValue()));

                    DefaultResponse<WorkPositionDetailResponse> workPositionDetailResponse = workPositionFeign.createWorkPosition("Bearer ".concat(loginResponse.getData().getToken()), workPositionRequest);
                    workPositionResponseMap.put(workPositionDetailResponse.getData().getWorkPosition().getDenomination(), workPositionDetailResponse.getData().getWorkPosition().getId());
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet cargos: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    public void migrateProfiles(MultipartFile file) {
        DefaultResponse<LoginResponse> loginResponse = login();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(4);
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
                    profileFeign.createProfile("Bearer ".concat(loginResponse.getData().getToken()), profileRequest);
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in sheet perfiles: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }

    private DefaultResponse<LoginResponse> login() {
        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(password);
        return loginFeign.login(loginRequest);
    }
}
