package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.DurationsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "durations", url="localhost:8080/core-api/v1/durations")
public interface DurationsFeign {

    @GetMapping
    DefaultResponse<List<DurationsResponse>> findAllDurations(@RequestHeader("Authorization") String token);
}
