package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodMaxDurationsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "workPeriodsMaxDurations", url="localhost:8080/core-api/v1/work-period-max-durations")
public interface WorkPeriodsMaxDurationsFeign {

    @GetMapping
    DefaultResponse<List<WorkPeriodMaxDurationsResponse>> findAllWorkPeriodsMaxDurations(@RequestHeader("Authorization") String token);
}
