package com.nelumbo.migration.feign;

import com.nelumbo.migration.exceptions.CustomErrorDecoder;
import com.nelumbo.migration.feign.dto.requests.WorkPeriodAssignRequest;
import com.nelumbo.migration.feign.dto.requests.WorkPeriodRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPeriodResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;

@FeignClient(name= "workPeriods", url="${hr-api}/work-periods", configuration = CustomErrorDecoder.class)
public interface WorkPeriodsFeign {

    @GetMapping("/findone-by-name")
    DefaultResponse<WorkPeriodResponse> findOneByName(@RequestHeader("Authorization") String token, @RequestParam String name);

    @PostMapping
    DefaultResponse<WorkPeriodResponse> createWorkPeriods(@RequestHeader("Authorization") String token, @RequestBody WorkPeriodRequest workPeriodRequest);

    @PostMapping(value = "/{workPeriodId}/work-period-assignments")
    void createWorkPeriodAssignments(@RequestHeader("Authorization") String token, @RequestBody WorkPeriodAssignRequest workPeriodAssignRequest, @PathVariable Long workPeriodId);
}
