package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.TabsRequest;
import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.TabsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "tabs", url="localhost:8080/core-api/v1/compensation-tabs")
public interface TabsFeign {

    @GetMapping("/simplified-search")
    DefaultResponse<List<TabsResponse>> simplifiedSearch(@RequestHeader("Authorization") String token, @RequestParam String search);

    @PostMapping
    void createTab(@RequestHeader("Authorization") String token, @RequestBody TabsRequest tabsRequest);
}
