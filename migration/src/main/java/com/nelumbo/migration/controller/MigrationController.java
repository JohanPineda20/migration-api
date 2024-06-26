package com.nelumbo.migration.controller;

import com.nelumbo.migration.service.MigrationService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
@RestController
@RequestMapping("/migration")
@RequiredArgsConstructor
public class MigrationController {

    private final MigrationService migrationService;
    @PostMapping
    public ResponseEntity<Void> migrateData(@RequestPart(value = "file") MultipartFile file){
        migrationService.migrateData(file);
        return new ResponseEntity<>(HttpStatus.ACCEPTED);
    }
}
