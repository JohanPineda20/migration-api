package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.StoreDetailRequest;
import com.nelumbo.migration.feign.dto.requests.StoreRequest;
import com.nelumbo.migration.feign.dto.requests.StoreWorkPeriodRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.StoreDetailResponse;
import com.nelumbo.migration.feign.dto.responses.StoreResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "store", url="localhost:8080/core-api/v1/stores")
public interface StoreFeign {
    @PostMapping
    DefaultResponse<StoreResponse> createStore(@RequestHeader("Authorization") String token,
                                               @RequestBody StoreRequest storeRequest);
    @PostMapping(path = "/{storeId}/details")
    DefaultResponse<StoreDetailResponse> createStoreDetails(@RequestHeader("Authorization") String token,
                                                            @RequestBody StoreDetailRequest storeDetailRequest,
                                                            @PathVariable Long storeId);
    @PostMapping(path = "/{storeId}/work-periods")
    void createStoreWorkPeriods(@RequestHeader("Authorization") String token,
                                @RequestBody StoreWorkPeriodRequest storeWorkPeriodRequest,
                                @PathVariable Long storeId);
}