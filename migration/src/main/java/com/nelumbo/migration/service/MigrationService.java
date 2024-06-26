package com.nelumbo.migration.service;

import com.nelumbo.migration.feign.CostCenterFeign;
import com.nelumbo.migration.feign.CountryFeign;
import com.nelumbo.migration.feign.LoginFeign;
import com.nelumbo.migration.feign.dto.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Slf4j
@Service
@RequiredArgsConstructor
public class MigrationService {
    private final CountryFeign countryFeign;
    private final CostCenterFeign costCenterFeign;
    private final LoginFeign loginFeign;
    @Value("${email}")
    private String email;
    @Value("${password}")
    private String password;

    public void migrateData(MultipartFile file) {
        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(password);
        DefaultResponse<LoginResponse> loginResponse = loginFeign.login(loginRequest);

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(0);
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < numberOfRows; i++) {
                try{
                    Row row = sheet.getRow(i);
                    CostCenterRequest costCenterRequest = new CostCenterRequest();
                    Cell cellCode = row.getCell(0);
                    costCenterRequest.setCode(cellCode.getCellType() == CellType.STRING ? cellCode.getStringCellValue() : "" + (int)cellCode.getNumericCellValue());
                    costCenterRequest.setDenomination(row.getCell(1).getStringCellValue());

                    DefaultResponse<List<CountryResponse>> countryResponse = countryFeign.findAll();
                    Long countryId = countryResponse.getData().stream()
                            .filter(country -> country.getName().equalsIgnoreCase(row.getCell(2).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElse(null);
                    DefaultResponse<List<CountryResponse>> stateResponse = countryFeign.findAllStatesByCountryId(countryId);
                    Long stateId = stateResponse.getData().stream()
                            .filter(state -> state.getName().equalsIgnoreCase(row.getCell(3).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElse(null);
                    DefaultResponse<List<CountryResponse>> cityResponse = countryFeign.findAllCitesByStateIdAndCountryId(countryId, stateId);
                    Long cityId = cityResponse.getData().stream()
                            .filter(city -> city.getName().equalsIgnoreCase(row.getCell(4).getStringCellValue()))
                            .findFirst().map(CountryResponse::getId).orElse(null);

                    costCenterRequest.setCountryId(countryId);
                    costCenterRequest.setStateId(stateId);
                    costCenterRequest.setCityId(cityId);
                    costCenterRequest.setStatusId(1L);

                    costCenterFeign.createCostCenter("Bearer ".concat(loginResponse.getData().getToken()), costCenterRequest);
                } catch (Exception e) {
                    log.error("Error processing row " + (i+1) + " in Excel: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
    }
}
