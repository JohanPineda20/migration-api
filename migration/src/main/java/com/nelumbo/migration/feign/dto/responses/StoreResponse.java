package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class StoreResponse {
    private Long id;
    private String code;
    private String denomination;
}
