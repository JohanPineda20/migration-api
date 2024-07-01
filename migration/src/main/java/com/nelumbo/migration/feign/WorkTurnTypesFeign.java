package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkTurnTypesResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestHeader;

import java.util.List;

@FeignClient(name= "workTurnTypes", url="localhost:8080/core-api/v1/work-turn-types")
public interface WorkTurnTypesFeign {

    @GetMapping
    DefaultResponse<List<WorkTurnTypesResponse>> findAllWorkTurnTypes(@RequestHeader("Authorization") String token);
}
