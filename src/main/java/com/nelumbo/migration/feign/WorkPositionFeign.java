package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.WorkPositionRequest;
import com.nelumbo.migration.feign.dto.requests.WorkPositionUpdateRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPositionDetailResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "work-positions", url="${hr-api}/work-positions")
public interface WorkPositionFeign {
    @PostMapping
    DefaultResponse<WorkPositionDetailResponse> createWorkPosition(@RequestHeader("Authorization") String token,
                                                                   @RequestBody WorkPositionRequest workPositionRequest);

    @PutMapping(path = "/{workPositionId}")
    void updateWorkPosition(@RequestHeader("Authorization") String token,
                                                                @RequestBody WorkPositionUpdateRequest workPositionRequest,
                                                                @PathVariable Long workPositionId);
}