package com.nelumbo.migration.feign.dto.responses;

import lombok.Data;

@Data
public class CountryResponse {
    private Long id;
    private String name;
}