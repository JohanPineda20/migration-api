package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class WorkPeriodRequest {

    private String name;
    private Long workPeriodTypeId;
    private List<WorkPeriodDetailRequest> workTurns;
    private Long workPeriodMaxDurationId;
    private Long workPeriodMaxDailyDurationId;
}
