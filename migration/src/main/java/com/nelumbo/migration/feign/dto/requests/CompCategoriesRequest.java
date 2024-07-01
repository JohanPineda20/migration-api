package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
public class CompCategoriesRequest {

    private String code;
    private String denomination;
    private Long statusId;
    private Map<String, Object> fieldsValues = new HashMap<>();
}
