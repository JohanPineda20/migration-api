package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.WorkPeriodRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

@FeignClient(name= "workPeriods", url="localhost:8080/core-api/v1/work-periods")
public interface WorkPeriodsFeign {

    @GetMapping("/findone-by-name")
    DefaultResponse<WorkPeriodResponse> findOneByName(@RequestHeader("Authorization") String token, @RequestParam String name);

    @PostMapping
    void createWorkPeriods(@RequestHeader("Authorization") String token, @RequestBody WorkPeriodRequest workPeriodRequest);
}
