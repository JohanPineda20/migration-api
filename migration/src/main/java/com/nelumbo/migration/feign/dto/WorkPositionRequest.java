package com.nelumbo.migration.feign.dto;

import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
public class WorkPositionRequest {
    private String code;
    private String denomination;
    private Map<String, Object> fieldsValues = new HashMap<>();
    private Long statusId;
    private Long authorizedStaff;
    private Long workPosCatId;
    private Long storeOrganizativeId;
    private Long storeId;
}
