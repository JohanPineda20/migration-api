package com.nelumbo.migration.feign.dto.requests;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class WorkPeriodDetailRequest {

    private String dateFrom;
    private String dateTo;
    private Integer dayOfWeek;
    private Long workTurnTypeId;
    private Long durationId;
}
