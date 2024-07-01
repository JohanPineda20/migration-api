package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DurationsResponse {
    private Long id;
    private String name;
    private int amount;
}
