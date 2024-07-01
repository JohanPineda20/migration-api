package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.*;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.http.MediaType;
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
}