package com.nelumbo.migration.feign.dto.requests;

import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import lombok.Getter;
import lombok.Setter;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
public class WorkPeriodRequest {

    private String name;
    private Long workPeriodTypeId;
    private List<WorkPeriodDetailRequest> workTurns;
    private Long workPeriodMaxDurationId;
    private Long workPeriodMaxDailyDurationId;
}
