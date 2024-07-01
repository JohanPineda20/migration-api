package com.nelumbo.migration.service;

import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Slf4j
@Service
@RequiredArgsConstructor
public class ServiceMigration {

    private static final String BEARER = "Bearer ";

    private final LoginFeign loginFeign;
    private final CompCategoriesFeign compCategoriesFeign;
    private final TabsFeign tabsFeign;
    private final WorksPositionCategoriesFeign worksPositionCategoriesFeign;
    private final WorkPeriodsFeign workPeriodsFeign;
    private final WorkPeriodsTypesFeign workPeriodsTypesFeign;
    private final WorkPeriodsMaxDurationsFeign workPeriodsMaxDurationsFeign;
    private final WorkPeriodsMaxDailyDurationsFeign workPeriodsMaxDailyDurationsFeign;
    private final DurationsFeign durationsFeign;
    private final WorkTurnTypesFeign workTurnTypesFeign;

    @Value("${email}")
    private String email;

    @Value("${password}")
    private String password;

    public File cargarCompensaciones(MultipartFile file) {

        String bearerToken = BEARER.concat(this.login());

        // Archivo modificado para devolver
        File modifiedFile = new File("modified_" + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            log.info("Entramos a recorrer el archivo");

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheetAt(0);

            log.info("Estamos con la hoja: " + sheet.getSheetName());

            // Cantidad de filas en la hoja
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            log.info("La cantidad de filas es: " + sheet.getPhysicalNumberOfRows());

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < numberOfRows; i++) {
                try {
                    // Sacamos el código y el nombre de la compensación
                    Row row = sheet.getRow(i);
                    String denomination = (row.getCell(1).getStringCellValue()).trim();
                    String code = (row.getCell(0).getCellType() == CellType.STRING ? row.getCell(0).getStringCellValue() : "" + (int) row.getCell(0).getNumericCellValue()).trim();

                    log.info("Compensacion a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por nombre
                    DefaultResponse<List<CompCategoriesResponse>> compCategoriesRes = compCategoriesFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsCompCategoriesByDeno = compCategoriesRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api: " + comp.getDenomination());
                                log.info("Compensacion del excel: " + denomination);
                                return comp.getDenomination().equalsIgnoreCase(denomination);
                            })
                            .findFirst().map(c -> true).orElse(false);

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsCompCategoriesByDeno ) {
                        log.info("Continuamos debido a que esa compensacion ya existe!");
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    compCategoriesRes = compCategoriesFeign.simplifiedSearch(bearerToken, code);
                    boolean existsCompCategoriesbyCode = compCategoriesRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api con codigo: " + comp.getCode());
                                log.info("Compensacion del excel con codigo: " + code);
                                return comp.getCode().equalsIgnoreCase(code);
                            }).findFirst().map(c -> true).orElse(false);

                    if (existsCompCategoriesbyCode && !existsCompCategoriesByDeno) {

                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a compensation-category with the code");
                        continue;
                    }

                    // Preparamos el objeto que irá en el body
                    CompCategoriesRequest compCategories = new CompCategoriesRequest();
                    compCategories.setCode(code);
                    compCategories.setDenomination(denomination);
                    compCategories.setStatusId(1L);

                    // Realizamos la petición
                    compCategoriesFeign.createCompensationCategories(bearerToken, compCategories);
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in Excel: " + e.getMessage());

                    // Agregar celda con el mensaje de error en la fila que falló
                    Row errorRow = sheet.getRow(i);
                    Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells() + 1);
                    errorCell.setCellValue("Error: " + e.getMessage());
                }
            }

            // Escribir el workbook modificado de nuevo en el archivo original
            try (FileOutputStream fileOut = new FileOutputStream(modifiedFile)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                log.error("Error writing modified Excel file: " + e.getMessage());
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return modifiedFile;
    }

    public File loadTabs(MultipartFile file) {

        String bearerToken = BEARER.concat(this.login());

        // Archivo modificado para devolver
        File modifiedFile = new File("modified_" + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            log.info("Entramos a recorrer el archivo");

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheetAt(0);

            log.info("Estamos con la hoja: " + sheet.getSheetName());

            // Cantidad de filas en la hoja
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            log.info("La cantidad de filas es: " + sheet.getPhysicalNumberOfRows());

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = workbook.createCellStyle();
            XSSFColor greenColor = new XSSFColor(Color.GREEN, null);
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(greenColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < numberOfRows; i++) {
                try {
                    // Sacamos el código y el nombre de la compensación
                    Row row = sheet.getRow(i);
                    String denomination = (row.getCell(1).getStringCellValue()).trim();
                    String code = (row.getCell(0).getCellType() == CellType.STRING ? row.getCell(0).getStringCellValue() : "" + (int) row.getCell(0).getNumericCellValue()).trim();
                    Long min_salary = (long) row.getCell(2).getNumericCellValue();
                    Long max_salary = (long) row.getCell(3).getNumericCellValue();

                    log.info("Tabulador a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por nombre
                    DefaultResponse<List<TabsResponse>> tabsRes = tabsFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsTabByDeno = tabsRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api: " + comp.getDenomination());
                                log.info("Compensacion del excel: " + denomination);
                                return comp.getDenomination().equalsIgnoreCase(denomination);
                            })
                            .findFirst().map(c -> true).orElse(false);

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsTabByDeno) {
                        log.info("Continuamos debido a que esa compensacion ya existe!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    tabsRes = tabsFeign.simplifiedSearch(bearerToken, code);
                    boolean existsTabByCode = tabsRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api con codigo: " + comp.getCode());
                                log.info("Compensacion del excel con codigo: " + code);
                                return comp.getCode().equalsIgnoreCase(code);
                            }).findFirst().map(c -> true).orElse(false);

                    if (existsTabByDeno && !existsTabByCode) {

                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a compensation-category with the code");
                        continue;
                    }

                    if (min_salary < 0 || max_salary < 0) {

                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: max_authorized_salary or min_authorized_salary can not be less than zero");
                        continue;
                    }

                    // Preparamos el objeto que irá en el body
                    TabsRequest tabsRequest = new TabsRequest();
                    tabsRequest.setCode(code);
                    tabsRequest.setDenomination(denomination);
                    tabsRequest.setMinAuthorizedSalary(min_salary);
                    tabsRequest.setMaxAuthorizedSalary(max_salary);
                    tabsRequest.setStatusId(1L);

                    // Realizamos la petición
                    tabsFeign.createTab(bearerToken, tabsRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in Excel: " + e.getMessage());
                    log.error("El tipo de excepction es: " + e.getClass().getSimpleName());

                    // Agregar celda con el mensaje de error en la fila que falló
                    Row errorRow = sheet.getRow(i);
                    Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells() + 1);
                    errorCell.setCellValue("Error: " + e.getMessage());
                }
            }

            // Escribir el workbook modificado de nuevo en el archivo original
            try (FileOutputStream fileOut = new FileOutputStream(modifiedFile)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                log.error("Error writing modified Excel file: " + e.getMessage());
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return modifiedFile;
    }

    public File loadWorkPositionCategories(MultipartFile file) {

        String bearerToken = BEARER.concat(this.login());

        // Archivo modificado para devolver
        File modifiedFile = new File("modified_" + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            log.info("Entramos a recorrer el archivo");

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheetAt(0);

            log.info("Estamos con la hoja: " + sheet.getSheetName());

            // Cantidad de filas en la hoja
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            log.info("La cantidad de filas es: " + sheet.getPhysicalNumberOfRows());

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = workbook.createCellStyle();
            XSSFColor greenColor = new XSSFColor(Color.GREEN, null);
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(greenColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < numberOfRows; i++) {
                try {
                    // Sacamos el código y el nombre de la compensación
                    Row row = sheet.getRow(i);
                    String denomination = (row.getCell(1).getStringCellValue()).trim();
                    String code = (row.getCell(0).getCellType() == CellType.STRING ? row.getCell(0).getStringCellValue() : "" + (int) row.getCell(0).getNumericCellValue()).trim();

                    log.info("Puesto a consultar con nombre: " + denomination + " \ncon codigo: " + code);

                    // Consultamos si existe la compensación por nombre
                    DefaultResponse<List<WorkPositionCategoryResponse>> worksPositionsCategoriesRes = worksPositionCategoriesFeign.simplifiedSearch(bearerToken, denomination);
                    boolean existsWorkByDeno = worksPositionsCategoriesRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api: " + comp.getDenomination());
                                log.info("Compensacion del excel: " + denomination);
                                return comp.getDenomination().equalsIgnoreCase(denomination);
                            })
                            .findFirst().map(c -> true).orElse(false);

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsWorkByDeno) {
                        log.info("Continuamos debido a que esa compensacion ya existe!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    // Consultamos si existe la compensación por el código (no debe existir dos compensaciones con el mismo código)
                    worksPositionsCategoriesRes = worksPositionCategoriesFeign.simplifiedSearch(bearerToken, code);
                    boolean existsWorkByCode = worksPositionsCategoriesRes.getData().stream()
                            .filter(comp -> {
                                log.info("Compensacion que viene del api con codigo: " + comp.getCode());
                                log.info("Compensacion del excel con codigo: " + code);
                                return comp.getCode().equalsIgnoreCase(code);
                            }).findFirst().map(c -> true).orElse(false);

                    if (existsWorkByCode && !existsWorkByDeno) {

                        // Agregar celda con el mensaje de error en la fila que falló
                        Row errorRow = sheet.getRow(i);
                        Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells());
                        errorCell.setCellValue("Error: exist a compensation-category with the code");
                        continue;
                    }

                    // Preparamos el objeto que irá en el body
                    WorkPositionCategoryRequest workPositionCategoryRequest = new WorkPositionCategoryRequest();
                    workPositionCategoryRequest.setCode(code);
                    workPositionCategoryRequest.setDenomination(denomination);
                    workPositionCategoryRequest.setStatusId(1L);

                    // Realizamos la petición
                    worksPositionCategoriesFeign.createWorkPositionCategory(bearerToken, workPositionCategoryRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in Excel: " + e.getMessage());

                    // Agregar celda con el mensaje de error en la fila que falló
                    Row errorRow = sheet.getRow(i);
                    Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells() + 1);
                    errorCell.setCellValue("Error: " + e.getMessage());
                }
            }

            // Escribir el workbook modificado de nuevo en el archivo original
            try (FileOutputStream fileOut = new FileOutputStream(modifiedFile)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                log.error("Error writing modified Excel file: " + e.getMessage());
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return modifiedFile;
    }

    public File loadWorkPeriods(MultipartFile file) {

        String bearerToken = BEARER.concat(this.login());

        List<WorkPeriodTypeResponse> workPeriodTypesList = workPeriodsTypesFeign.findAllWorkPeriodTypes(bearerToken).getData();
        List<WorkPeriodMaxDurationsResponse> workPeriodMaxDurationsList = workPeriodsMaxDurationsFeign.findAllWorkPeriodsMaxDurations(bearerToken).getData();
        List<WorkPeriodMaxDailyDurationsResponse> workPeriodMaxDailyDurationsResponseList = workPeriodsMaxDailyDurationsFeign.findAllWorkPeriodsMaxDailyDurations(bearerToken).getData();
        List<DurationsResponse> durations = durationsFeign.findAllDurations(bearerToken).getData();
        List<WorkTurnTypesResponse> workturntypes = workTurnTypesFeign.findAllWorkTurnTypes(bearerToken).getData();

        // Archivo modificado para devolver
        File modifiedFile = new File("modified_" + file.getOriginalFilename());

        // Para abrir el workbook y que se cierre automáticamente al finalizar
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {

            log.info("Entramos a recorrer el archivo");

            // Nos posicionamos en la primera hoja
            Sheet sheet = workbook.getSheetAt(0);

            log.info("Estamos con la hoja: " + sheet.getSheetName());

            // Cantidad de filas en la hoja
            int numberOfRows = sheet.getPhysicalNumberOfRows();

            log.info("La cantidad de filas es: " + sheet.getPhysicalNumberOfRows());

            // Crear un estilo de celda con color verde para los datos insertados correctamente
            CellStyle cellStyle = workbook.createCellStyle();
            XSSFColor greenColor = new XSSFColor(Color.GREEN, null);
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(greenColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
            for (int i = 1; i < numberOfRows; i++) {
                try {

                    List<WorkPeriodDetailRequest> workPeriodDetailList = new ArrayList<>();

                    Row row = sheet.getRow(i);

                    String name = (row.getCell(0).getStringCellValue()).trim();
                    String periodType = (row.getCell(1).getStringCellValue()).trim();
                    String keywordMaxDuration = (row.getCell(2).getStringCellValue()).trim();
                    Integer maxDailyDuration = (row.getCell(3) == null) ? null : (int)row.getCell(3).getNumericCellValue();

                    log.info("Periodo de trabajo a consultar con nombre: " + name);

                    // Consultamos si existe el periodo de trabajo por nombre
                    DefaultResponse<WorkPeriodResponse> workPeriodResponse = workPeriodsFeign.findOneByName(bearerToken, name);
                    boolean existsWorkPeriod = workPeriodResponse.getData() != null &&
                            workPeriodResponse.getData().getName().equalsIgnoreCase(name);

                    // Si existe, seguimos a la siguiente para no volverla a insertar
                    if (existsWorkPeriod) {
                        log.info("Continuamos debido a que existe una jornada laboral con ese nombre!");
                        row.getCell(0).setCellStyle(cellStyle);
                        continue;
                    }

                    Long idWorkPeriodType;
                    Long idWorkPeriodMaxDailyDuration = null;
                    if(periodType.equalsIgnoreCase("Horario Fijo")) {

                        idWorkPeriodType = workPeriodTypesList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("Fixed Scheduled/Regular shift"))
                                .map(wpt -> wpt.getId()).findFirst().get();

                        log.info("El valor de max daily duration es: " + maxDailyDuration);

                        if(maxDailyDuration == 24) {

                            idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpmd -> wpmd.getDuration() == 24)
                                    .map(wpmd -> wpmd.getId()).findFirst().get();

                        } else if(maxDailyDuration == 12) {

                            idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpt -> wpt.getDuration() == 12)
                                    .map(wpt -> wpt.getId()).findFirst().get();

                        } else if(maxDailyDuration == 8) {

                            idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpt -> wpt.getDuration() == 8)
                                    .map(wpt -> wpt.getId()).findFirst().get();

                        } else if(maxDailyDuration == 6) {

                            idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpt -> wpt.getDuration() == 6)
                                    .map(wpt -> wpt.getId()).findFirst().get();

                        } else if(maxDailyDuration == 4) {

                            idWorkPeriodMaxDailyDuration = workPeriodMaxDailyDurationsResponseList
                                    .stream()
                                    .filter(wpt -> wpt.getDuration() == 4)
                                    .map(wpt -> wpt.getId()).findFirst().get();

                        } else {
                            throw new Exception("There is no max_daily_duration");
                        }

                    } else if(periodType.equalsIgnoreCase("Frecuencia Variable")) {

                        idWorkPeriodType = workPeriodTypesList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("Variable frecuency shift"))
                                .map(wpt -> wpt.getId()).findFirst().get();

                    } else {
                        throw new Exception("There is no period with that name");
                    }

                    Long idWorkPeriodMaxDuration;
                    if(keywordMaxDuration.equalsIgnoreCase("48HWFT")) {

                        idWorkPeriodMaxDuration = workPeriodMaxDurationsList
                                .stream()
                                .filter(wpmd -> wpmd.getKeyword().equalsIgnoreCase("48HWFT"))
                                .map(wpmd -> wpmd.getId()).findFirst().get();

                    } else if(keywordMaxDuration.equalsIgnoreCase("40HWFT")) {

                        idWorkPeriodMaxDuration = workPeriodMaxDurationsList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("40HWFT"))
                                .map(wpt -> wpt.getId()).findFirst().get();

                    } else if(keywordMaxDuration.equalsIgnoreCase("30HWPT")) {

                        idWorkPeriodMaxDuration = workPeriodMaxDurationsList
                                .stream()
                                .filter(wpmd -> wpmd.getKeyword().equalsIgnoreCase("30HWPT"))
                                .map(wpmd -> wpmd.getId()).findFirst().get();

                    } else if(keywordMaxDuration.equalsIgnoreCase("48HWST")) {

                        idWorkPeriodMaxDuration = workPeriodMaxDurationsList
                                .stream()
                                .filter(wpt -> wpt.getName().equalsIgnoreCase("48HWST"))
                                .map(wpt -> wpt.getId()).findFirst().get();

                    } else {
                        throw new Exception("There is no max_duration with that keyword");
                    }

                    // Nos posicionamos en la segunda hoja donde estan los tunos de trabajo
                    Sheet workTurnsSheet = workbook.getSheetAt(1);

                    log.info("Estamos con la hoja: " + workTurnsSheet.getSheetName());

                    // Cantidad de filas en la hoja
                    int numberOfWorkTurns = workTurnsSheet.getPhysicalNumberOfRows();

                    log.info("La cantidad de filas de turnos de trabajo: " + workTurnsSheet.getPhysicalNumberOfRows());

                    // Crear un estilo de celda con color verde para los datos insertados correctamente
//                    CellStyle cellStyle = workbook.createCellStyle();
//                    XSSFColor greenColor = new XSSFColor(Color.GREEN, null);
//                    ((XSSFCellStyle) cellStyle).setFillForegroundColor(greenColor);
//                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                    // Recorrer la cantidad de filas a partir de la posición 1 porque la 0 son los nombres de las columnas
                    for (int j = 1; j < numberOfWorkTurns; j++) {

                        Row workTurnRow = workTurnsSheet.getRow(j);
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
                            dateFrom = workTurnRow.getCell(1).getDateCellValue();
                            dateTo = workTurnRow.getCell(2).getDateCellValue();
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
                    workPeriodsFeign.createWorkPeriods(bearerToken, workPeriodRequest);
                    row.getCell(0).setCellStyle(cellStyle);
                } catch (Exception e) {
                    log.error("Error processing row " + (i + 1) + " in Excel: " + e.getMessage());

                    // Agregar celda con el mensaje de error en la fila que falló
                    Row errorRow = sheet.getRow(i);
                    Cell errorCell = errorRow.createCell(errorRow.getPhysicalNumberOfCells() + 1);
                    errorCell.setCellValue("Error: " + e.getMessage());
                }
            }

            // Escribir el workbook modificado de nuevo en el archivo original
            try (FileOutputStream fileOut = new FileOutputStream(modifiedFile)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                log.error("Error writing modified Excel file: " + e.getMessage());
            }
        } catch (Exception e) {
            log.error("Error processing Excel file: " + e.getMessage());
        }
        return modifiedFile;
    }

    private String login() {
        // Realizamos el login para obtener un token
        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(password);
        return loginFeign.login(loginRequest).getData().getToken();
    }
}
