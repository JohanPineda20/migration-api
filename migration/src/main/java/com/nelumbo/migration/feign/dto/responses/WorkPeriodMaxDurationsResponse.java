package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class WorkPeriodMaxDurationsResponse {
    private Long id;
    private String name;
    private int duration;
    private String keyword;
}
