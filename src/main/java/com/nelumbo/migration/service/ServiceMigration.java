package com.nelumbo.migration.service;

import com.nelumbo.migration.feign.*;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
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
    private static final String MODIFIED = "modified_";

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


}
