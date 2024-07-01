package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.*;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "work-positions", url="localhost:8080/core-api/v1/work-positions")
public interface WorkPositionFeign {
    @PostMapping
    DefaultResponse<WorkPositionDetailResponse> createWorkPosition(@RequestHeader("Authorization") String token,
                                                                   @RequestBody WorkPositionRequest workPositionRequest);
}