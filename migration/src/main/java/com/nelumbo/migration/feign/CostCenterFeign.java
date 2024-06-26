package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.CostCenterRequest;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "costCenters", url="localhost:8080/core-api/v1/cost-centers")
public interface CostCenterFeign {
    @PostMapping
    void createCostCenter(@RequestHeader("Authorization") String token, @RequestBody CostCenterRequest costCenterRequest);
}