package com.nelumbo.migration.feign.dto.responses;

import lombok.Data;

import java.util.Map;

@Data
public class ModelFieldsResponse {
    private Long id;
    private Long profileSectionId;
    private String name;
    private Map<String, Object> validations;
    private String placeholder;
    private String label;
    private Long catalogueId;
    private FieldTypeResponse fieldType;
    private boolean visible;
    private boolean used;
    private int position;
    private boolean locked;
}
