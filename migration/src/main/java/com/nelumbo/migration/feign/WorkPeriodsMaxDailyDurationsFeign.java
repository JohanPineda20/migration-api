package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodMaxDailyDurationsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "workPeriodsMaxDailyDurations", url="localhost:8080/core-api/v1/work-period-max-daily-durations")
public interface WorkPeriodsMaxDailyDurationsFeign {

    @GetMapping
    DefaultResponse<List<WorkPeriodMaxDailyDurationsResponse>> findAllWorkPeriodsMaxDailyDurations(@RequestHeader("Authorization") String token);
}
