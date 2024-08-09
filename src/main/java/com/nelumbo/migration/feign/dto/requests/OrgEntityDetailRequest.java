package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
public class OrgEntityDetailRequest {
    private String name;
    private Map<String, Object> fieldValues = new HashMap<>();
    private Long parentId;
}
