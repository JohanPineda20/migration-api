package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodTypeResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "workPeriodsTypes", url="localhost:8080/core-api/v1/work-period-types")
public interface WorkPeriodsTypesFeign {

    @GetMapping
    DefaultResponse<List<WorkPeriodTypeResponse>> findAllWorkPeriodTypes(@RequestHeader("Authorization") String token);
}
