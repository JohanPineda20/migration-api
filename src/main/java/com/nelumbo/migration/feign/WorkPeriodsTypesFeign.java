package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodTypeResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "workPeriodsTypes", url="${hr-api}/work-period-types")
public interface WorkPeriodsTypesFeign {

    @GetMapping
    DefaultResponse<List<WorkPeriodTypeResponse>> findAllWorkPeriodTypes(@RequestHeader("Authorization") String token);
}
