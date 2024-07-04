package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.CostCenterRequest;
import com.nelumbo.migration.feign.dto.responses.CostCenterResponse;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "costCenters", url="${hr-api}/cost-centers")
public interface CostCenterFeign {
    @PostMapping
    DefaultResponse<CostCenterResponse> createCostCenter(@RequestHeader("Authorization") String token, @RequestBody CostCenterRequest costCenterRequest);
}