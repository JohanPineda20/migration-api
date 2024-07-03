package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.Map;
@Getter
@Setter
public class StoreRequest {
    private String code;
    private String denomination;
    private Long countryId;
    private Long stateId;
    private Long cityId;
    private Map<String, Object> fieldsValues = new HashMap<>();
    private Long statusId;
    private String address;
    private String zipcode;
    private Long georefDistance;
    private Double latitude;
    private Double longitude;
    private Long costCenterId;
}
