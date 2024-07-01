package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
public class CostCenterRequest {
    private String code;
    private String denomination;
    private Long countryId;
    private Long stateId;
    private Long cityId;
    private Map<String, Object> fieldsValues = new HashMap<>();
    private Long statusId;
}
